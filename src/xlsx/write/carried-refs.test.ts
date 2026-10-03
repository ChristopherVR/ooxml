import { describe, expect, it } from 'vitest';
import { SourceIndex } from '../read/package.js';
import { createWorkbook } from '../workbook.js';
import {
	type CarriedRefEdits,
	carriedRefEdits,
	patchCarriedFormula,
	patchCarriedPart,
	patchCarriedParts,
	workbookRefEdits,
} from './carried-refs.js';
import { PackageWriter } from './package-writer.js';

const none: CarriedRefEdits = { shifts: [], deleted: [], renamed: [] };
const insertRows = (sheet: string, at: number, count: number): CarriedRefEdits => ({
	...none,
	shifts: [{ sheet, shift: { axis: 'row', at, count } }],
});

// As Excel writes them (xl2.xlsx from the review generator gen2.ps1).
const PIVOT =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
	'<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" ' +
	'recordCount="6"><cacheSource type="worksheet"><worksheetSource ref="A1:C7" sheet="Data"/>' +
	'</cacheSource><cacheFields count="0"/></pivotCacheDefinition>';
const CHART =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
	'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart>' +
	'<c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>' +
	'<c:cat><c:strRef><c:f>Data!$A$2:$A$7</c:f></c:strRef></c:cat>' +
	'<c:val><c:numRef><c:f>Data!$B$2:$B$7</c:f><c:numCache/></c:numRef></c:val>' +
	'<c:extLst><c:ext uri="x"><c15:datalabelsRange xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">' +
	'<c15:f>Other!$C$2:$C$7</c15:f></c15:datalabelsRange></c:ext></c:extLst>' +
	'</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>';

describe('patchCarriedPart', () => {
	it('returns undefined without edits or for unknown parts', () => {
		expect(patchCarriedPart('xl/pivotCache/pivotCacheDefinition1.xml', PIVOT, none)).toBe(
			undefined,
		);
		const edits = insertRows('Data', 0, 2);
		expect(patchCarriedPart('xl/custom.xml', '<root ref="A1"/>', edits)).toBe(undefined);
		expect(patchCarriedPart('xl/media/image1.png', CHART, edits)).toBe(undefined);
		expect(patchCarriedPart('xl/charts/chart1.xml', CHART, insertRows('Data', 100, 1))).toBe(
			undefined,
		);
	});

	it('shifts a pivot cache source like Excel (rows inserted above it)', () => {
		const out = patchCarriedPart('p.xml', PIVOT, insertRows('Data', 0, 2)) ?? '';
		expect(out).toContain('<worksheetSource ref="A3:C9" sheet="Data"/>');
		expect(out).toContain('recordCount="6"');
	});

	it('renames the pivot source sheet, escaping the attribute', () => {
		const edits: CarriedRefEdits = { ...none, renamed: [['Data', 'R&D "x"']] };
		const out = patchCarriedPart('p.xml', PIVOT, edits) ?? '';
		expect(out).toContain('<worksheetSource ref="A1:C7" sheet="R&amp;D &quot;x&quot;"/>');
	});

	it('keeps named, external and deleted-sheet pivot sources, and a fully deleted ref', () => {
		const named = PIVOT.replace('ref="A1:C7" sheet="Data"', 'name="Sales"');
		expect(patchCarriedPart('p.xml', named, { ...none, renamed: [['Data', 'D']] })).toBe(undefined);
		expect(patchCarriedPart('p.xml', PIVOT, { ...none, deleted: ['Data'] })).toBe(undefined);
		const gone: CarriedRefEdits = {
			...none,
			shifts: [{ sheet: 'Data', shift: { axis: 'row', at: 0, count: -10 } }],
		};
		expect(patchCarriedPart('p.xml', PIVOT, gone)).toBe(undefined);
		const other = insertRows('Elsewhere', 0, 2);
		expect(patchCarriedPart('p.xml', PIVOT, other)).toBe(undefined);
	});

	it('shifts and renames chart formulas, including extension ones', () => {
		const edits: CarriedRefEdits = {
			shifts: [{ sheet: 'Data', shift: { axis: 'row', at: 0, count: 2 } }],
			deleted: [],
			renamed: [
				['Data', 'My Data'],
				['Other', 'O'],
			],
		};
		const out = patchCarriedPart('xl/charts/chart2.xml', CHART, edits) ?? '';
		expect(out).toContain("<c:f>'My Data'!$B$3</c:f>");
		expect(out).toContain("<c:f>'My Data'!$A$4:$A$9</c:f>");
		expect(out).toContain("<c:f>'My Data'!$B$4:$B$9</c:f>");
		expect(out).toContain('<c15:f>O!$C$2:$C$7</c15:f>');
		expect(out).toContain('<c:numCache/>');
	});

	it('turns references to a deleted sheet into #REF! as Excel does', () => {
		const out = patchCarriedPart('c.xml', CHART, { ...none, deleted: ['Data'] }) ?? '';
		expect(out).toContain('<c:f>#REF!</c:f>');
		expect(out).toContain('<c15:f>Other!$C$2:$C$7</c15:f>');
	});
});

describe('patchCarriedFormula', () => {
	it('swaps two sheet names without merging them', () => {
		const edits: CarriedRefEdits = {
			...none,
			renamed: [
				['A', 'B'],
				['B', 'A'],
			],
		};
		expect(patchCarriedFormula('(A!$A$1,B!$B$2)', '', edits)).toBe('(B!$A$1,A!$B$2)');
	});

	it('shifts only inside a band for insert-cells edits', () => {
		const edits: CarriedRefEdits = {
			...none,
			shifts: [{ sheet: 'S', shift: { axis: 'row', at: 1, count: 1, band: { lo: 0, hi: 0 } } }],
		};
		expect(patchCarriedFormula('S!$A$2:$A$3', '', edits)).toBe('S!$A$3:$A$4');
		expect(patchCarriedFormula('S!$B$2', '', edits)).toBe('S!$B$2');
	});
});

describe('carriedRefEdits', () => {
	it('matches sheets by sheetId and keys shifts by source name', () => {
		const shift = { axis: 'col' as const, at: 0, count: 1 };
		const current = [
			{ name: 'Renamed', sheetId: 1, shifts: [shift] },
			{ name: 'New', sheetId: 9, shifts: [shift] },
		];
		const edits = carriedRefEdits(
			[
				{ name: 'Data', sheetId: 1 },
				{ name: 'Gone', sheetId: 2 },
			],
			current,
			(sheet) => sheet.shifts,
		);
		expect(edits).toEqual({
			shifts: [{ sheet: 'Data', shift }],
			deleted: ['Gone'],
			renamed: [['Data', 'Renamed']],
		});
	});
});

const enc = (text: string): Uint8Array => new TextEncoder().encode(text);

function sourcePackage(): SourceIndex {
	const parts = new Map<string, Uint8Array>([
		[
			'[Content_Types].xml',
			enc(
				'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
					'<Default Extension="xml" ContentType="application/xml"/></Types>',
			),
		],
		[
			'_rels/.rels',
			enc(
				'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
			),
		],
		[
			'xl/workbook.xml',
			enc(
				'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
					'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
					'<sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/>' +
					'</sheets></workbook>',
			),
		],
		['xl/pivotCache/pivotCacheDefinition1.xml', enc(PIVOT)],
		['xl/charts/chart1.xml', enc(CHART)],
		['xl/charts/chart2.xml', enc(CHART)],
	]);
	return new SourceIndex(parts);
}

describe('patchCarriedParts', () => {
	it('rewrites only parts still carried byte for byte', async () => {
		const source = sourcePackage();
		const workbook = createWorkbook();
		workbook.sheets[0] = { ...workbook.sheets[0]!, name: 'Data 2', sheetId: 1 };
		const edits = workbookRefEdits(workbook, source);
		expect(edits).toEqual({ shifts: [], deleted: ['Other'], renamed: [['Data', 'Data 2']] });

		const writer = new PackageWriter(source);
		writer.carry('xl/pivotCache/pivotCacheDefinition1.xml');
		writer.carry('xl/charts/chart1.xml');
		writer.carry('xl/charts/chart2.xml');
		writer.add('xl/charts/chart2.xml', '<c:chartSpace>modelled</c:chartSpace>');
		patchCarriedParts(writer, edits);
		expect(writer.carriedParts().map(([name]) => name)).toEqual([]);

		const JSZip = (await import('jszip')).default;
		const zip = await JSZip.loadAsync(await writer.build());
		const text = (name: string) => zip.file(name)?.async('string') ?? Promise.resolve('');
		expect(await text('xl/pivotCache/pivotCacheDefinition1.xml')).toContain('sheet="Data 2"');
		const chart = await text('xl/charts/chart1.xml');
		expect(chart).toContain("<c:f>'Data 2'!$B$1</c:f>");
		expect(chart).toContain('<c15:f>#REF!</c15:f>');
		expect(await text('xl/charts/chart2.xml')).toBe('<c:chartSpace>modelled</c:chartSpace>');
	});
});
