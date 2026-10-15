import { mkdtemp, readdir, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createFileOperations } from './node';

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
