import { execSync } from 'node:child_process';
import { mkdir, mkdtemp, readdir, realpath, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { vi } from 'vitest';
import { createFileOperations } from './node';
import { isUnsafePathSegment, pathSegments } from './node-paths';

async function withRoot(run: (root: string, outside: string) => Promise<void>) {
	const root = await mkdtemp(join(tmpdir(), 'ooxml-root-'));
	const outside = await mkdtemp(join(tmpdir(), 'ooxml-outside-'));
	try {
		await run(root, outside);
	} finally {
		await rm(root, { recursive: true, force: true });
		await rm(outside, { recursive: true, force: true });
	}
}

async function link(target: string, path: string): Promise<boolean> {
	try {
		await symlink(target, path, process.platform === 'win32' ? 'junction' : 'dir');
		return true;
	} catch {
		return false;
	}
}

describe('file operations confinement', () => {
	test('rejects reads and writes through a link that leaves the root', async () => {
		await withRoot(async (root, outside) => {
			await writeFile(join(outside, 'secret.pptx'), new Uint8Array([1]));
			if (!(await link(outside, join(root, 'escape')))) return;
			const files = createFileOperations(root);
			await expect(files.readFile('escape/secret.pptx', ['.pptx'])).rejects.toThrow('outside');
			await expect(
				files.writeFile('escape/new.pptx', ['.pptx'], new Uint8Array([2])),
			).rejects.toThrow('outside');
			await expect(files.directory('escape/sub')).rejects.toThrow('outside');
			expect((await readdir(outside)).sort()).toEqual(['secret.pptx']);
		});
	});

	test('rejects files over the size cap', async () => {
		await withRoot(async (root) => {
			await writeFile(join(root, 'big.pptx'), new Uint8Array(64 * 1024 * 1024 + 1));
			const files = createFileOperations(root);
			await expect(files.readFile('big.pptx', ['.pptx'])).rejects.toThrow('64 MiB');
		});
	});

	test('returns root-relative paths and hides the host path in errors', async () => {
		await withRoot(async (root) => {
			const files = createFileOperations(root);
			const created = await files.create('a.docx', ['.docx'], async () => new Uint8Array([0]));
			expect(created.filePath).toBe('a.docx');
			const edited = await files.edit('a.docx', ['.docx'], async () => ({
				bytes: new Uint8Array([1]),
			}));
			expect(edited.filePath).toBe('a.docx');
			const error = await files
				.inspect('missing-dir/x.docx', ['.docx'], async () => null)
				.catch((e: Error) => e);
			expect((error as Error).message).not.toContain(root);
			expect(await files.directory('out/media')).toBe('out/media');
		});
	});
});

describe('lexical path refusal', () => {
	const REFUSED = 'Path is outside the configured root';
	const B = '\\';
	const windowsOnly = [
		'C:x.docx',
		'C:../x.docx',
		`..${B}outside${B}a.docx`,
		'C:/elsewhere/x.docx',
		`C:${B}elsewhere${B}x.docx`,
		'plain.bat:s.docx',
		'a.docx:s',
		'a.docx::$DATA',
		'sub:s/a.docx',
		'NUL.docx',
		'nul',
		'con.docx',
		'COM1.docx',
		'lpt¹.docx',
		'CONIN$.docx',
		'sub/aux.docx',
		'a.docx.',
		'a.docx ',
		'.. /a.docx',
		'a?.docx',
		'a*.docx',
	];
	const refused = [
		...(process.platform === 'win32' ? windowsOnly : []),
		`${B}${B}127.0.0.1${B}C$${B}Windows${B}x.docx`,
		`${B}${B}127.0.0.1${B}C$${B}missing${B}x.docx`,
		`${B}${B}host${B}share${B}x.docx`,
		'//host/share/x.docx',
		`${B}${B}?${B}C:${B}x.docx`,
		`${B}${B}.${B}C:${B}x.docx`,
		`${B}${B}.${B}pipe${B}x.docx`,
		'//./pipe/x.docx',
		'//?/UNC/host/share/x.docx',
		`${B}rooted${B}x.docx`,
		'../outside-exists/a.docx',
		'../outside-missing/a.docx',
		'sub/../../a.docx',
		'a\u0000.docx',
	];

	test.each(refused)('refuses %j with one error before any filesystem call', async (path) => {
		// The root does not exist, so any filesystem call would fail with ENOENT instead.
		const files = createFileOperations(join(tmpdir(), 'ooxml-no-such-root', 'root'));
		const read = vi.fn(async () => null);
		await expect(files.inspect(path, ['.docx'], read)).rejects.toThrow(new Error(REFUSED));
		await expect(files.create(path, ['.docx'], async () => new Uint8Array([1]))).rejects.toThrow(
			new Error(REFUSED),
		);
		await expect(files.directory(path)).rejects.toThrow(new Error(REFUSED));
		expect(read).not.toHaveBeenCalled();
	});

	test.runIf(process.platform === 'win32')('an alternate stream name creates nothing', async () => {
		await withRoot(async (root) => {
			const files = createFileOperations(root);
			await expect(
				files.create('plain.bat:s.docx', ['.docx'], async () => new Uint8Array([1])),
			).rejects.toThrow(REFUSED);
			expect(await readdir(root)).toEqual([]);
		});
	});

	test('names are refused only where the platform gives them another meaning', () => {
		for (const name of ['a: b.docx', 'aux', 'v1.', 'why?.docx', 'NUL.docx', 'trailing ']) {
			expect(isUnsafePathSegment(name, 'win32')).toBe(true);
			expect(isUnsafePathSegment(name, 'linux')).toBe(false);
		}
		expect(isUnsafePathSegment('a\u0000b', 'linux')).toBe(true);
	});

	test('a backslash separates path segments only on Windows', () => {
		expect(pathSegments('a\\b/c', 'win32')).toEqual(['a', 'b', 'c']);
		expect(pathSegments('a\\b/c', 'linux')).toEqual(['a\\b', 'c']);
	});

	test.runIf(process.platform !== 'win32')('POSIX keeps a backslash in a name', async () => {
		await withRoot(async (root) => {
			const files = createFileOperations(root);
			await writeFile(join(root, 'back\\slash.docx'), new Uint8Array([1]));
			expect([...(await files.readFile('back\\slash.docx', ['.docx']))]).toEqual([1]);
			await expect(files.readFile('..\\slash.docx', ['.docx'])).rejects.toThrow();
		});
	});

	test.runIf(process.platform !== 'win32')('POSIX keeps names Windows refuses', async () => {
		await withRoot(async (root) => {
			const files = createFileOperations(root);
			for (const name of ['a: b.docx', 'aux/a.docx', 'v1./a.docx', 'why?.docx']) {
				await mkdir(join(root, name, '..'), { recursive: true });
				await writeFile(join(root, name), new Uint8Array([1]));
				expect([...(await files.readFile(name, ['.docx']))]).toEqual([1]);
			}
		});
	});

	test('an absolute path may reach the root through a link', async () => {
		await withRoot(async (root, outside) => {
			await writeFile(join(root, 'a.docx'), new Uint8Array([1]));
			const spelled = join(outside, 'link-to-root');
			if (!(await link(root, spelled))) return;
			const viaReal = createFileOperations(root);
			const viaLink = createFileOperations(spelled);
			const real = join(await realpath(root), 'a.docx');
			expect([...(await viaReal.readFile(join(spelled, 'a.docx'), ['.docx']))]).toEqual([1]);
			expect([...(await viaLink.readFile(real, ['.docx']))]).toEqual([1]);
			expect([...(await viaLink.readFile(join(spelled, 'a.docx'), ['.docx']))]).toEqual([1]);
			expect(await viaReal.directory(spelled)).toBe('.');
			const created = await viaReal.create(join(spelled, 'new.docx'), ['.docx'], async () =>
				Uint8Array.of(2),
			);
			expect(created.filePath).toBe('new.docx');
		});
	});

	test('an absolute link spelling cannot leave the root', async () => {
		await withRoot(async (root, outside) => {
			await writeFile(join(outside, 'secret.docx'), new Uint8Array([1]));
			if (!(await link(outside, join(root, 'escape')))) return;
			const files = createFileOperations(root);
			for (const path of [
				join(root, 'escape', 'secret.docx'),
				join(outside, 'secret.docx'),
				join(outside, 'missing', 'x.docx'),
			])
				await expect(files.readFile(path, ['.docx'])).rejects.toThrow(new Error(REFUSED));
		});
	});

	test.runIf(process.platform === 'win32')('a short path name reaches the same root', async () => {
		await withRoot(async (root) => {
			await writeFile(join(root, 'a.docx'), new Uint8Array([1]));
			const short = execSync(`cmd /c for %I in ("${root}") do @echo %~sI`, {
				encoding: 'utf8',
			}).trim();
			const files = createFileOperations(root);
			expect([...(await files.readFile(join(short, 'a.docx'), ['.docx']))]).toEqual([1]);
			const fromShort = createFileOperations(short);
			expect([...(await fromShort.readFile(join(root, 'a.docx'), ['.docx']))]).toEqual([1]);
		});
	});

	test('an absolute path outside the root is refused like any other', async () => {
		await withRoot(async (root, outside) => {
			await writeFile(join(outside, 'a.docx'), new Uint8Array([1]));
			const files = createFileOperations(root);
			await expect(files.readFile(join(outside, 'a.docx'), ['.docx'])).rejects.toThrow(
				new Error(REFUSED),
			);
		});
	});

	test('paths that stay inside the root keep working', async () => {
		await withRoot(async (root) => {
			await mkdir(join(root, 'sub'));
			await writeFile(join(root, 'sub', 'a.docx'), new Uint8Array([1]));
			const files = createFileOperations(root);
			for (const path of [
				'sub/a.docx',
				join('sub', 'a.docx'),
				'./sub/./a.docx',
				'sub/../sub/a.docx',
				join(root, 'sub', 'a.docx'),
			])
				expect([...(await files.readFile(path, ['.docx']))]).toEqual([1]);
			expect(await files.directory('.')).toBe('.');
			expect(await files.directory(join(root, 'sub'))).toBe('sub');
		});
	});

	test('a missing folder inside the root is still reported as missing', async () => {
		await withRoot(async (root) => {
			const files = createFileOperations(root);
			await expect(files.readFile('missing/a.docx', ['.docx'])).rejects.toMatchObject({
				code: 'ENOENT',
			});
		});
	});

	test('the extension is checked before the filesystem', async () => {
		const files = createFileOperations(join(tmpdir(), 'ooxml-no-such-root', 'root'));
		await expect(files.readFile('a.exe', ['.docx'])).rejects.toThrow('Unsupported file extension');
	});
});
