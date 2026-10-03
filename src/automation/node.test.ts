import { mkdtemp, readFile, rm, writeFile, symlink, mkdir, readdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { createFileOperations } from './node.js';

test('file operations scope paths, preserve originals on failure and serialize concurrent edits', async () => {
	const root = await mkdtemp(join(tmpdir(), 'ooxml-files-'));
	try {
		const files = createFileOperations(root);
		await files.create('sample.docx', ['.docx'], async () => new Uint8Array([0]));
		await expect(
			files.create('sample.docx', ['.docx'], async () => new Uint8Array([9])),
		).rejects.toThrow();
		await expect(files.inspect('../escape.docx', ['.docx'], async () => null)).rejects.toThrow(
			'outside',
		);
		await expect(files.inspect('sample.exe', ['.docx'], async () => null)).rejects.toThrow(
			'extension',
		);
		await expect(
			files.edit('sample.docx', ['.docx'], async () => {
				throw new Error('invalid edit');
			}),
		).rejects.toThrow('invalid edit');
		expect([...(await readFile(join(root, 'sample.docx')))]).toEqual([0]);
		await Promise.all(
			Array.from({ length: 4 }, () =>
				files.edit('sample.docx', ['.docx'], async (bytes) => ({
					bytes: new Uint8Array([bytes[0]! + 1]),
				})),
			),
		);
		expect([...(await readFile(join(root, 'sample.docx')))]).toEqual([4]);
		await writeFile(join(root, 'target.docx'), new Uint8Array([8]));
		await expect(
			files.edit(
				'sample.docx',
				['.docx'],
				async () => ({ bytes: new Uint8Array([5]) }),
				'target.docx',
			),
		).rejects.toThrow();
		expect([...(await readFile(join(root, 'target.docx')))]).toEqual([8]);
		expect((await readdir(root)).some((file) => file.endsWith('.tmp'))).toBe(false);
		await mkdir(join(root, 'nested'));
		await symlink(tmpdir(), join(root, 'nested', 'outside'), 'junction');
		await expect(
			files.inspect('nested/outside/escape.docx', ['.docx'], async () => null),
		).rejects.toThrow('outside');
	} finally {
		await rm(root, { recursive: true, force: true });
	}
});
