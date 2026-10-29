import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { headerFooterPicture } from '../layout/index';
import { saveXlsx } from '../write/index';
import { loadXlsx } from './index';

const fixture = () =>
	new Uint8Array(
		readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', 'excel-header-images.xlsx')),
	);

it('reads the header and footer pictures of Excel with their sizes', async () => {
	const wb = await loadXlsx(fixture());
	const sheet = wb.sheets[0]!;
	expect(Object.keys(sheet.headerFooterImages ?? {}).sort()).toEqual([
		'CF',
		'LH',
		'LHFIRST',
		'RFEVEN',
		'RH',
	]);
	expect(sheet.headerFooterImages?.LH).toMatchObject({
		partName: 'xl/media/image1.png',
		contentType: 'image/png',
		widthPt: 30,
		heightPt: 15,
		title: 'hf-logo',
	});
	// Page 1 shows the first-page header, page 2 the even-page footer.
	expect(headerFooterPicture(sheet, 'header', 'left', 1, 1)?.partName).toBe('xl/media/image1.png');
	expect(headerFooterPicture(sheet, 'header', 'right', 1, 1)).toBeUndefined();
	expect(headerFooterPicture(sheet, 'header', 'right', 2, 3)?.partName).toBe('xl/media/image1.png');
	expect(headerFooterPicture(sheet, 'footer', 'right', 2, 2)?.partName).toBe('xl/media/image1.png');
});

it('keeps the header picture parts on save', async () => {
	const wb = await loadXlsx(fixture());
	const zip = await JSZip.loadAsync(await saveXlsx(wb));
	expect(Object.keys(zip.files).some((n) => n.startsWith('xl/media/'))).toBe(true);
	expect(await zip.file('xl/worksheets/sheet1.xml')!.async('string')).toContain('legacyDrawingHF');
	const again = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	expect(Object.keys(again.sheets[0]!.headerFooterImages ?? {})).toHaveLength(5);
});
