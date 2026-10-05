import { link, open, realpath, rename, unlink } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, relative, resolve } from 'node:path';
import { randomUUID } from 'node:crypto';

export interface ByteEditResult {
	bytes: Uint8Array;
}

/** Local filesystem boundary used by the repository-owned MCP adapters. */
export function createFileOperations(rootDir = process.cwd()) {
	const pending = new Map<string, Promise<unknown>>();
	const maxBytes = 64 * 1024 * 1024;
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
	return {
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
				return { filePath: path, saved: true };
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
				return { ...result, filePath: output, saved: true };
			});
		},
	};
}
