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

/** Best of three runs, in milliseconds. */
async function timed(run: () => Promise<unknown>): Promise<number> {
	let best = Number.POSITIVE_INFINITY;
	for (let attempt = 0; attempt < 3; attempt++) {
		const start = performance.now();
		await run();
		best = Math.min(best, performance.now() - start);
	}
	return best;
}

const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A one-sheet workbook whose header drawing is `vml`. */
async function workbookWithHeaderDrawing(vml: string): Promise<Uint8Array> {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
	);
	zip.file(
		'xl/workbook.xml',
		`<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
	);
	zip.file(
		'xl/_rels/workbook.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
	);
	zip.file(
		'xl/worksheets/sheet1.xml',
		`<worksheet xmlns="${S}" xmlns:r="${R}"><sheetData/><legacyDrawingHF r:id="rId9"/></worksheet>`,
	);
	zip.file(
		'xl/worksheets/_rels/sheet1.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R}/vmlDrawing" Target="../drawings/v.vml"/></Relationships>`,
	);
	zip.file('xl/drawings/v.vml', vml);
	return zip.generateAsync({ type: 'uint8array' });
}

it('reads a header drawing in time linear in its shapes, empty ones included', async () => {
	const run = async (count: number) => {
		const bytes = await workbookWithHeaderDrawing(
			`<xml xmlns:v="urn:schemas-microsoft-com:vml">${'<v:shape/>'.repeat(count)}</xml>`,
		);
		return () => loadXlsx(bytes);
	};
	const small = await run(20_000);
	const large = await run(80_000);
	expect((await timed(large)) / (await timed(small))).toBeLessThan(8);
}, 60_000);
