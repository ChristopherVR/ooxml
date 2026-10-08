import JSZip from 'jszip';
import { describe, it, expect } from 'vitest';
import { listEmbeddedPackages, replaceEmbeddedPackage } from './embedded-packages';
import { createXlsx, setXlsxCells, readXlsxRange } from '../automation/xlsx';

describe('embedded package save-back', () => {
	it('edits a nested workbook and preserves unrelated parent parts', async () => {
		const zip = new JSZip();
		zip.file('word/embeddings/budget.xlsx', await createXlsx(['Budget']));
		zip.file('word/document.xml', '<original-formatting/>');
		zip.file('word/embeddings/legacy.bin', new Uint8Array([1, 2, 3]));
		const original = await zip.generateAsync({ type: 'uint8array' });
		const parts = await listEmbeddedPackages(original);
		expect(parts[1]?.kind).toBe('unsupported');
		const workbook = parts[0]!;
		const edited = await setXlsxCells(workbook.bytes, 0, [{ address: 'B2', input: '42' }]);
		const saved = await replaceEmbeddedPackage(original, workbook.path, edited.bytes);
		const reopened = await JSZip.loadAsync(saved);
		expect(await reopened.file('word/document.xml')?.async('string')).toBe(
			'<original-formatting/>',
		);
		expect(await reopened.file('word/embeddings/legacy.bin')?.async('uint8array')).toEqual(
			new Uint8Array([1, 2, 3]),
		);
		expect(
			(await readXlsxRange((await listEmbeddedPackages(saved))[0]!.bytes, 0, 'B2')).cells[0]?.value,
		).toBe(42);
	});
	it('rejects missing, binary and traversal targets', async () => {
		const zip = await new JSZip().generateAsync({ type: 'uint8array' });
		for (const path of [
			'word/embeddings/missing.xlsx',
			'word/embeddings/object.bin',
			'../secret.xlsx',
		])
			await expect(replaceEmbeddedPackage(zip, path, zip)).rejects.toThrow();
	});
});
