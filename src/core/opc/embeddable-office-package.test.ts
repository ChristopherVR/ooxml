import JSZip from 'jszip';
import { expect, it } from 'vitest';

import { readEmbeddableOfficePackage } from './embeddable-office-package';

it('admits an Office package only as the corresponding object class', async () => {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', '<Types/>');
	zip.file('xl/workbook.xml', '<workbook/>');
	const bytes = await zip.generateAsync({ type: 'uint8array' });

	await expect(readEmbeddableOfficePackage(bytes, 'XLSX')).resolves.toEqual({
		extension: 'xlsx',
		type: 'excel',
		progId: 'Excel.Sheet.12',
		mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	});
	await expect(readEmbeddableOfficePackage(bytes, 'docx')).rejects.toThrow('file type');
	await expect(readEmbeddableOfficePackage(bytes, 'xlsm')).rejects.toThrow('Unsupported object');
	await expect(readEmbeddableOfficePackage(new Uint8Array([1, 2, 3]), 'xlsx')).rejects.toThrow();
});
