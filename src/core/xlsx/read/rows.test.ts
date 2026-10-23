import { readFileSync } from 'node:fs';
import path from 'node:path';
import { deflateRawSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { deflateZeros, lyingZipOf } from '../../opc/__fixtures__/lying-zip';
import { NS } from '../../xml/index';
import { putCell } from '../cells';
import type { Workbook } from '../model';
import { numberFormatView } from '../numfmt/view';
import { styleAt } from '../styles';
import { createWorkbook, createWorksheet } from '../workbook';
import { saveXlsx } from '../write/index';
import { loadXlsx } from './load';
import { readXlsxRows, type ReadXlsxRowsOptions, type XlsxRow } from './rows';

const fixture = (name: string) =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', name)));

async function rowsOf(bytes: Uint8Array, options?: ReadXlsxRowsOptions): Promise<XlsxRow[]> {
	const rows: XlsxRow[] = [];
	for await (const row of readXlsxRows(bytes, options)) rows.push(row);
	return rows;
}

/** The rows `loadXlsx` gives for a sheet, in the reader's shape. */
function expected(workbook: Workbook, index: number): XlsxRow[] {
	const sheet = workbook.sheets[index];
	const rows: XlsxRow[] = [];
	for (const [row, cells] of [...(sheet?.rows ?? [])].sort((a, b) => a[0] - b[0])) {
		const out = [...cells]
			.sort((a, b) => a[0] - b[0])
			.filter(([, cell]) => cell.value !== null || cell.formula)
			.map(([col, cell]) => ({
				col,
				value: cell.value,
				...(cell.formula ? { formula: cell.formula } : {}),
				...numberFormatView(styleAt(workbook, cell.styleId).numFmt, cell.value, workbook.date1904),
			}));
		if (out.length) rows.push({ row, cells: out });
	}
	return rows;
}

/** A workbook whose first sheet part is `sheetData` (with the namespaces `root` declares). */
async function withSheetData(sheetData: string, root = `xmlns="${NS.x}"`): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(await saveXlsx(createWorkbook()));
	zip.file(
		'xl/worksheets/sheet1.xml',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet ${root}><sheetViews><sheetView workbookViewId="0"/></sheetViews>${sheetData}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
	);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/** `count` rows of five cells: numbers, a shared formula, inline text, a boolean and an error. */
function manyRows(count: number, each: (row: number) => string = () => ''): string {
	const out = ['<sheetData>'];
	for (let r = 1; r <= count; r++) {
		const formula =
			r === 1 ? '<f t="shared" ref="B1:B99999" si="0">A1*2</f>' : '<f t="shared" si="0"/>';
		out.push(
			`<row r="${r}"><c r="A${r}"><v>${r / 4}</v></c><c r="B${r}">${formula}<v>${r / 2}</v></c>`,
			`<c r="C${r}" t="inlineStr"><is><t>text ${r} &amp; more</t></is></c>`,
			`<c r="D${r}" t="b"><v>${r % 2}</v></c><c r="E${r}" t="e"><v>#N/A</v></c>${each(r)}</row>`,
		);
	}
	out.push('</sheetData>');
	return out.join('');
}

describe('readXlsxRows', () => {
	it.each([
		'excel-features.xlsx',
		'excel-1904.xlsx',
		'openpyxl-features.xlsx',
		'openpyxl-styles.xlsx',
		'openpyxl-1904.xlsx',
		'excel-sparklines.xlsx',
	])('gives the cells of every sheet of %s as loadXlsx does', async (name) => {
		const bytes = fixture(name);
		const workbook = await loadXlsx(bytes);
		for (const [index, sheet] of workbook.sheets.entries()) {
			const rows = await rowsOf(bytes, { sheet: index });
			expect(rows, sheet.name).toEqual(expected(workbook, index));
			expect(await rowsOf(bytes, { sheet: sheet.name.toUpperCase() })).toEqual(rows);
		}
	});

	it('reports dates and number formats as the automation reads do', async () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		const dated = { ...styleAt(workbook, 0), numFmt: 'yyyy-mm-dd' };
		workbook.styles.push(dated, { ...styleAt(workbook, 0), numFmt: '0.00%' });
		putCell(sheet, 0, 0, { value: 45352, styleId: 1 });
		putCell(sheet, 0, 1, { value: 0.25, styleId: 2 });
		putCell(sheet, 0, 2, { value: 'label' });
		putCell(sheet, 1, 0, { value: null, styleId: 1 });
		const rows = await rowsOf(await saveXlsx(workbook));
		expect(rows).toEqual([
			{
				row: 0,
				cells: [
					{ col: 0, value: 45352, numberFormat: 'yyyy-mm-dd', iso: '2024-03-01' },
					{ col: 1, value: 0.25, numberFormat: '0.00%' },
					{ col: 2, value: 'label' },
				],
			},
		]);
	});

	it('reads a sheet larger than many inflated chunks, row by row, as loadXlsx does', async () => {
		const bytes = await withSheetData(manyRows(8000));
		const rows = await rowsOf(bytes);
		expect(rows).toHaveLength(8000);
		expect(rows).toEqual(expected(await loadXlsx(bytes), 0));
		expect(rows[7999]?.cells[1]).toEqual({ col: 1, value: 4000, formula: 'A8000*2' });
	});

	it('reads prefixed elements (Open XML SDK output) as loadXlsx does', async () => {
		const bytes = await withSheetData(
			manyRows(3000).replace(/<(\/?)(sheetData|row|c|v|f|is|t)\b/g, '<$1x:$2'),
			`xmlns:x="${NS.x}"`,
		);
		const rows = await rowsOf(bytes);
		expect(rows).toHaveLength(3000);
		expect(rows).toEqual(expected(await loadXlsx(bytes), 0));
	});

	it.each([
		['a comment holding a row start', (r: number) => (r === 1500 ? '<!-- <row r="9"> -->' : '')],
		[
			'CDATA holding the end of the cell data',
			(r: number) =>
				r === 1500 ? `<c r="F${r}" t="inlineStr"><is><t><![CDATA[</sheetData>]]></t></is></c>` : '',
		],
		['a processing instruction', (r: number) => (r === 1500 ? '<?keep?>' : '')],
		['a namespace declaration', (r: number) => (r === 1500 ? '<y:c xmlns:y="urn:y"/>' : '')],
	])('goes on from the same row when a piece has %s', async (_, each) => {
		const bytes = await withSheetData(manyRows(3000, each));
		const rows = await rowsOf(bytes);
		expect(rows).toHaveLength(3000);
		expect(rows).toEqual(expected(await loadXlsx(bytes), 0));
	});

	it('stops reading when the loop stops, before markup further on is seen', async () => {
		// Broken after the first rows: loading the workbook fails, the first rows can be read.
		const bytes = await withSheetData(
			manyRows(3000).replace('<row r="2999">', '<row r="2999"><c>'),
		);
		await expect(loadXlsx(bytes)).rejects.toThrow();
		const first: number[] = [];
		for await (const { row } of readXlsxRows(bytes)) {
			first.push(row);
			if (first.length === 100) break;
		}
		expect(first).toEqual(Array.from({ length: 100 }, (_, i) => i));
		await expect(rowsOf(bytes)).rejects.toThrow();
	});

	it('picks a sheet by position or name, and refuses others', async () => {
		const workbook = createWorkbook();
		workbook.sheets.push(createWorksheet('Second', 2));
		putCell(workbook.sheets[1]!, 4, 2, { value: 'here' });
		const bytes = await saveXlsx(workbook);
		const second = [{ row: 4, cells: [{ col: 2, value: 'here' }] }];
		expect(await rowsOf(bytes, { sheet: 1 })).toEqual(second);
		expect(await rowsOf(bytes, { sheet: 'second' })).toEqual(second);
		await expect(rowsOf(bytes, { sheet: 2 })).rejects.toThrow(RangeError);
		await expect(rowsOf(bytes, { sheet: -1 })).rejects.toThrow(RangeError);
		await expect(rowsOf(bytes, { sheet: 'Third' })).rejects.toThrow(/no sheet named "Third"/);
		const zip = await JSZip.loadAsync(bytes);
		const rels = 'xl/_rels/workbook.xml.rels';
		const xml = await zip.file(rels)!.async('string');
		const chart = xml.replace(/<Relationship [^>]*sheet2\.xml[^>]*>/, (tag) =>
			tag.replace('/worksheet"', '/chartsheet"'),
		);
		expect(chart).not.toBe(xml);
		zip.file(rels, chart);
		const charted = await zip.generateAsync({ type: 'uint8array' });
		await expect(rowsOf(charted, { sheet: 1 })).rejects.toThrow(/is a chartsheet/);
	});

	it('reports what loadXlsx warns about, once each', async () => {
		const bytes = await withSheetData(
			'<sheetData><row r="1"><c r="A1" t="s"><v>99</v></c><c r="B1" t="s"><v>98</v></c></row></sheetData>',
		);
		const warnings: string[] = [];
		await rowsOf(bytes, { onWarning: (message) => warnings.push(message) });
		expect(warnings).toEqual((await loadXlsx(bytes)).warnings);
		expect(warnings).toEqual(['A cell refers to a missing shared string.']);
	});

	it('applies the load limits while inflating', async () => {
		const bytes = await withSheetData(manyRows(2000));
		await expect(rowsOf(bytes, { limits: { maxPartBytes: 100_000 } })).rejects.toMatchObject({
			name: 'WorkbookTooLargeError',
			limit: 'maxPartBytes',
		});
		await expect(rowsOf(bytes, { limits: { maxInputBytes: 1000 } })).rejects.toMatchObject({
			limit: 'maxInputBytes',
		});
		// A sheet part that claims 100 bytes and inflates to 64 MiB stops at the first chunk past it.
		const zip = await JSZip.loadAsync(bytes);
		const entries = await Promise.all(
			Object.values(zip.files)
				.filter((entry) => !entry.dir && entry.name !== 'xl/worksheets/sheet1.xml')
				.map(async (entry) => {
					const data = await entry.async('nodebuffer');
					return { name: entry.name, data: deflateRawSync(data), declared: data.length };
				}),
		);
		const bomb = lyingZipOf([
			...entries,
			{ name: 'xl/worksheets/sheet1.xml', data: await deflateZeros(64 << 20), declared: 100 },
		]);
		const started = Date.now();
		await expect(rowsOf(bomb)).rejects.toThrow(/inflates past its declared size/);
		expect(Date.now() - started).toBeLessThan(5000);
	});
});
