import { link, open, realpath, rename, unlink } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, relative, resolve, sep } from 'node:path';
import { randomUUID } from 'node:crypto';

export interface ByteEditResult {
	bytes: Uint8Array;
}

/** Local filesystem boundary used by the repository-owned MCP adapters. */
export function createFileOperations(rootDir = process.cwd()) {
	const pending = new Map<string, Promise<unknown>>();
	const maxBytes = 64 * 1024 * 1024;
	/** Root-relative, forward-slash form of a confined absolute path (never the host path). */
	async function rel(target: string) {
		const root = await realpath(resolve(rootDir));
		return relative(root, target).split(sep).join('/') || '.';
	}
	/** Replace host paths in an error message so callers only see root-relative text. */
	async function scrub<T>(action: () => Promise<T>): Promise<T> {
		try {
			return await action();
		} catch (error) {
			if (!(error instanceof Error)) throw error;
			const hosts = [resolve(rootDir), await realpath(resolve(rootDir)).catch(() => '')];
			let message = error.message;
			for (const host of hosts.filter(Boolean).sort((a, b) => b.length - a.length))
				message = message
					.split(host + sep)
					.join('')
					.split(host)
					.join('<root>');
			const safe = new Error(message, { cause: undefined });
			(safe as NodeJS.ErrnoException).code = (error as NodeJS.ErrnoException).code;
			throw safe;
		}
	}
	async function pathFor(file: string, extensions: readonly string[]) {
		const root = await realpath(resolve(rootDir));
		const path = resolve(root, file);
		const parent = await realpath(dirname(path));
		let actual = resolve(parent, basename(path));
		const inside = (target: string) => {
			const rel = relative(root, target);
			if (rel === '..' || rel.startsWith('../') || rel.startsWith('..\\') || isAbsolute(rel))
				throw new Error('Path is outside the configured root');
		};
		inside(actual);
		try {
			actual = await realpath(actual);
			inside(actual);
		} catch (error) {
			if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
		}
		if (!extensions.includes(extname(actual).toLowerCase()))
			throw new Error('Unsupported file extension');
		return actual;
	}
	async function read(path: string) {
		const handle = await open(path, 'r');
		try {
			if ((await handle.stat()).size > maxBytes) throw new Error('File exceeds 64 MiB');
			const bytes = await handle.readFile();
			if (bytes.length > maxBytes) throw new Error('File exceeds 64 MiB');
			return new Uint8Array(bytes);
		} finally {
			await handle.close();
		}
	}
	async function write(path: string, bytes: Uint8Array, replace: boolean) {
		if (bytes.length > maxBytes) throw new Error('Output exceeds 64 MiB');
		const temp = `${path}.${randomUUID()}.tmp`;
		const handle = await open(temp, 'wx', 0o600);
		try {
			await handle.writeFile(bytes);
			await handle.sync();
		} catch (error) {
			await handle.close();
			await unlink(temp).catch(() => {});
			throw error;
		}
		await handle.close();
		try {
			if (replace) await rename(temp, path);
			else await link(temp, path);
		} finally {
			await unlink(temp).catch(() => {});
		}
	}
	async function serial<T>(path: string, action: () => Promise<T>): Promise<T> {
		if (process.platform === 'win32') path = path.toLowerCase();
		const before = pending.get(path) ?? Promise.resolve();
		const current = before.catch(() => {}).then(action);
		pending.set(path, current);
		try {
			return await current;
		} finally {
			if (pending.get(path) === current) pending.delete(path);
		}
	}
	/** Confine a directory (it may not exist yet) and return its root-relative path. */
	async function directoryFor(dir: string) {
		const root = await realpath(resolve(rootDir));
		let probe = resolve(root, dir);
		const rest: string[] = [];
		for (;;) {
			try {
				probe = await realpath(probe);
				break;
			} catch (error) {
				if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
				const parent = dirname(probe);
				if (parent === probe) throw error;
				rest.unshift(basename(probe));
				probe = parent;
			}
		}
		const actual = resolve(probe, ...rest);
		const r = relative(root, actual);
		if (r === '..' || r.startsWith('../') || r.startsWith('..\\') || isAbsolute(r))
			throw new Error('Path is outside the configured root');
		return r.split(sep).join('/') || '.';
	}
	const api = {
		async inspect<T>(
			file: string,
			extensions: readonly string[],
			inspect: (bytes: Uint8Array) => Promise<T>,
		) {
			const path = await pathFor(file, extensions);
			return serial(path, async () => inspect(await read(path)));
		},
		async create(file: string, extensions: readonly string[], create: () => Promise<Uint8Array>) {
			const path = await pathFor(file, extensions);
			return serial(path, async () => {
				await write(path, await create(), false);
				return { filePath: await rel(path), saved: true };
			});
		},
		async edit<T extends ByteEditResult>(
			file: string,
			extensions: readonly string[],
			edit: (bytes: Uint8Array) => Promise<T>,
			outputPath?: string,
		) {
			const path = await pathFor(file, extensions);
			const output = outputPath ? await pathFor(outputPath, extensions) : path;
			if (extname(output).toLowerCase() !== extname(path).toLowerCase())
				throw new Error('Save-as must keep the source file extension');
			return serial(path, async () => {
				const { bytes, ...result } = await edit(await read(path));
				await write(output, bytes, output === path);
				return { ...result, filePath: await rel(output), saved: true };
			});
		},
		/** Read a confined file (no serialisation; use inside {@link exclusive}). */
		async readFile(file: string, extensions: readonly string[]) {
			return read(await pathFor(file, extensions));
		},
		/** Atomically replace a confined file (no serialisation; use inside {@link exclusive}). */
		async writeFile(file: string, extensions: readonly string[], bytes: Uint8Array) {
			await write(await pathFor(file, extensions), bytes, true);
		},
		/** Run `action` holding the per-path lock of a confined file. */
		async exclusive<T>(file: string, extensions: readonly string[], action: () => Promise<T>) {
			return serial(await pathFor(file, extensions), action);
		},
		/** Confine a directory under the root and return its root-relative path. */
		directory: directoryFor,
		/** Root-relative description of a root-confined file path. */
		relativePath: async (file: string, extensions: readonly string[]) =>
			rel(await pathFor(file, extensions)),
	};
	type Api = typeof api;
	const wrapped = {} as Record<string, unknown>;
	for (const [name, fn] of Object.entries(api))
		wrapped[name] = (...args: unknown[]) =>
			scrub(() => (fn as (...a: unknown[]) => Promise<unknown>)(...args));
	return wrapped as unknown as Api;
}
