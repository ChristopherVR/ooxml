import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { writeChartWorkbookUpdates } from './chart-xlsx-writer';

async function workbook(
	cell = '<c r="A1" s="4" t="inlineStr"><is><t>Old</t></is></c>',
	second = '',
) {
	const zip = new JSZip();
	zip.file(
		'xl/workbook.xml',
		'<workbook><sheets><sheet name="Data" r:id="r1"/></sheets></workbook>',
	);
	zip.file(
		'xl/_rels/workbook.xml.rels',
		'<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>',
	);
	zip.file(
		'xl/worksheets/sheet1.xml',
		`<worksheet><sheetData><row r="1">${cell}${second}</row></sheetData><extLst><ext uri="retained"/></extLst></worksheet>`,
	);
	zip.file('unrelated.xml', '<preserve value="unchanged"/>');
	return zip.generateAsync({ type: 'uint8array' });
}

const write = { formula: 'Data!$A$1', isNumeric: false, values: ['New'] };

describe('atomic chart workbook label transactions', () => {
	it('changes only the addressed cells and preserves worksheet formatting and unrelated parts', async () => {
		const result = await writeChartWorkbookUpdates(await workbook(), [write], { atomic: true });
		expect(result.unresolved).toBe(0);
		const zip = await JSZip.loadAsync(result.bytes!);
		const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
		expect(xml).toContain('s="4"');
		expect(xml).toContain('New');
		expect(xml).toContain('uri="retained"');
		expect(await zip.file('unrelated.xml')!.async('string')).toBe('<preserve value="unchanged"/>');
	});
	it('rejects a mixed resolvable and missing range without returning partially changed bytes', async () => {
		const bytes = await workbook();
		await expect(
			writeChartWorkbookUpdates(bytes, [write, { ...write, formula: 'Missing!A1' }], {
				atomic: true,
			}),
		).rejects.toThrow('source range');
		expect(
			await (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml')!.async('string'),
		).toContain('Old');
	});
	it('rejects formulas even outside the edited range to avoid stale dependent results', async () => {
		await expect(
			writeChartWorkbookUpdates(
				await workbook(undefined, '<c r="B1"><f>A1</f><v>Old</v></c>'),
				[write],
				{ atomic: true },
			),
		).rejects.toThrow('formula-dependent');
	});
	it('rejects conflicting overlapping writes, but accepts duplicate identical references', async () => {
		await expect(
			writeChartWorkbookUpdates(await workbook(), [write, { ...write, values: ['Conflict'] }], {
				atomic: true,
			}),
		).rejects.toThrow('conflicting');
		expect(
			(await writeChartWorkbookUpdates(await workbook(), [write, write], { atomic: true }))
				.unresolved,
		).toBe(0);
	});
	it('rejects rich worksheet strings rather than stripping their run formatting', async () => {
		await expect(
			writeChartWorkbookUpdates(
				await workbook('<c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>Old</t></r></is></c>'),
				[write],
				{ atomic: true },
			),
		).rejects.toThrow('rich worksheet');
	});
	it('rejects phonetic and unknown worksheet string children rather than discarding them', async () => {
		await expect(
			writeChartWorkbookUpdates(
				await workbook(
					'<c r="A1" t="inlineStr"><is><t>Old</t><rPh sb="0" eb="3"><t>Reading</t></rPh></is></c>',
				),
				[write],
				{ atomic: true },
			),
		).rejects.toThrow('rich worksheet');
	});
	it('rejects corrupt workbooks and worksheets without cell data', async () => {
		await expect(
			writeChartWorkbookUpdates(new Uint8Array([1]), [write], { atomic: true }),
		).rejects.toThrow('readable');
		const zip = await JSZip.loadAsync(await workbook());
		zip.file('xl/worksheets/sheet1.xml', '<worksheet/>');
		await expect(
			writeChartWorkbookUpdates(await zip.generateAsync({ type: 'uint8array' }), [write], {
				atomic: true,
			}),
		).rejects.toThrow('no cell data');
	});
});
