import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseAddress } from '../address.js';
import { getCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { styleAt } from '../styles.js';
import { loadXlsx } from './index.js';

const fixture = (name: string) =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', name)));
const at = (wb: Workbook, sheet: number, ref: string) => {
	const address = parseAddress(ref);
	if (!address) throw new Error(ref);
	return getCell(wb.sheets[sheet]!, address.row, address.col);
};

describe('loadXlsx with an Excel-authored workbook', () => {
	it('reads sheets, visibility, defined names and document properties', async () => {
		const wb = await loadXlsx(fixture('excel-features.xlsx'));
		expect(wb.format).toBe('xlsx');
		expect(wb.sheets.map((s) => [s.name, s.state])).toEqual([
			['Main', 'visible'],
			['Second', 'visible'],
			['Secret', 'hidden'],
		]);
		expect(wb.definedNames).toEqual([{ name: 'Rate', formula: 'Main!$C$2' }]);
		expect(wb.sheets[0]!.pageSetup?.printArea).toEqual({
			start: { row: 0, col: 0 },
			end: { row: 11, col: 4 },
		});
		expect(wb.properties).toMatchObject({
			title: 'Excel fixture',
			creator: 'Excel generator',
			application: 'Microsoft Excel',
		});
		expect(wb.warnings).toEqual([]);
	});

	it('expands shared formulas per cell and keeps cached values', async () => {
		const wb = await loadXlsx(fixture('excel-features.xlsx'));
		expect(at(wb, 0, 'D2')).toMatchObject({ formula: 'B2*C2', value: 3.75 });
		expect(at(wb, 0, 'D6')).toMatchObject({ formula: 'B6*C6', value: 93.75 });
		expect(at(wb, 0, 'E4')?.formula).toBe('D4/SUM($D$2:$D$6)');
		expect(at(wb, 0, 'C8')).toMatchObject({
			formula: 'MAX(B2:B6*C2:C6)',
			arrayRange: { start: { row: 7, col: 2 }, end: { row: 7, col: 2 } },
		});
		expect(at(wb, 0, 'D8')).toMatchObject({ formula: 'IF(B8>10,"big","small")', value: 'big' });
		expect(at(wb, 0, 'D12')).toMatchObject({ formula: '1/0', value: { error: '#DIV/0!' } });
		expect(at(wb, 0, 'C12')?.value).toBe(true);
	});

	it('reads rich text runs and _xHHHH_ escapes', async () => {
		const wb = await loadXlsx(fixture('excel-features.xlsx'));
		const rich = at(wb, 0, 'A10');
		expect(rich?.value).toBe('Rich text');
		expect(rich?.richText?.map((r) => r.text)).toEqual(['Rich', ' ', 'text']);
		expect(rich?.richText?.[2]?.font?.color).toEqual({ rgb: 'FFFF0000' });
		expect(at(wb, 0, 'A11')?.value).toBe('Line 1\r\nLine 2');
	});

	it('maps cell styles to resolved, deduplicated formats', async () => {
		const wb = await loadXlsx(fixture('excel-features.xlsx'));
		const header = styleAt(wb, at(wb, 0, 'A1')?.styleId);
		expect(header.font).toMatchObject({ bold: true, color: { theme: 0 } });
		expect(header.fill).toEqual({
			type: 'pattern',
			pattern: 'solid',
			fgColor: { rgb: 'FF0070C0' },
		});
		expect(header.border.bottom).toEqual({ style: 'medium', color: { indexed: 64 } });
		expect(styleAt(wb, at(wb, 0, 'E2')?.styleId).numFmt).toBe('0.0%');
		expect(styleAt(wb, at(wb, 0, 'B12')?.styleId).numFmt).toBe('#,##0.00_);[Red](#,##0.00)');
		expect(styleAt(wb, at(wb, 0, 'B2')?.styleId).fill).toMatchObject({
			fgColor: { theme: 5, tint: 0.59999389629810485 },
		});
		expect(styleAt(wb, at(wb, 0, 'H2')?.styleId).cellStyleName).toBe('Hyperlink');
		expect(wb.namedStyles.map((s) => s.name)).toEqual(['Normal', 'Hyperlink']);
		expect(wb.styles[0]?.cellStyleName).toBeUndefined();
	});

	it('reads the theme palette in SpreadsheetML index order (lt1 first)', async () => {
		const wb = await loadXlsx(fixture('excel-features.xlsx'));
		expect(wb.theme.colors[0]).toBe('FFFFFF');
		expect(wb.theme.colors[1]).toBe('000000');
		expect(wb.theme.colors).toHaveLength(12);
	});

	it('reads views, columns, rows and merges', async () => {
		const sheet = (await loadXlsx(fixture('excel-features.xlsx'))).sheets[0]!;
		expect(sheet.view).toMatchObject({
			zoom: 90,
			freeze: { rows: 1, cols: 1 },
			topLeft: { row: 1, col: 1 },
		});
		expect(sheet.tabColor).toEqual({ rgb: 'FF50B000' });
		expect(sheet.columns[0]).toMatchObject({ min: 0, max: 0, customWidth: true });
		expect(sheet.columns[2]).toMatchObject({ min: 6, max: 7 });
		expect(sheet.rowInfo.get(10)).toEqual({ height: 30, customHeight: true });
		expect(sheet.rowInfo.get(12)).toEqual({ hidden: true });
		expect(sheet.merges).toEqual([{ start: { row: 13, col: 0 }, end: { row: 14, col: 2 } }]);
	});

	it('reads conditional formats with their dxf styles and the x14 link', async () => {
		const sheet = (await loadXlsx(fixture('excel-features.xlsx'))).sheets[0]!;
		const rules = sheet.conditionalFormats.flatMap((f) => f.rules);
		expect(rules.map((r) => r.type)).toEqual(['cellIs', 'dataBar', 'colorScale', 'iconSet']);
		expect(rules[0]).toMatchObject({
			operator: 'greaterThan',
			formulas: ['6'],
			style: { fill: { pattern: 'solid', fgColor: { rgb: 'FFFFC79C' } } },
		});
		expect(rules[1]).toMatchObject({ extensionId: expect.stringMatching(/^\{[0-9A-F-]+\}$/) });
		expect(sheet.preserved.get('extLst')?.[0]).toContain('x14:conditionalFormattings');
	});

	it('reads validations, hyperlinks and threaded comments', async () => {
		const sheet = (await loadXlsx(fixture('excel-features.xlsx'))).sheets[0]!;
		expect(sheet.dataValidations[0]).toMatchObject({
			type: 'list',
			formula1: '"Yes,No"',
			showDropDown: true,
			prompt: 'Yes or No',
		});
		expect(sheet.dataValidations[1]).toMatchObject({
			type: 'whole',
			formula1: '1',
			formula2: '10',
		});
		expect(sheet.hyperlinks).toEqual([
			{
				range: { start: { row: 1, col: 7 }, end: { row: 1, col: 7 } },
				target: 'https://example.org/',
				tooltip: 'Example tip',
			},
			{
				range: { start: { row: 2, col: 7 }, end: { row: 2, col: 7 } },
				location: "'Second'!A1",
				display: 'Go to Second',
			},
		]);
		expect(sheet.comments).toEqual([
			{ address: { row: 1, col: 0 }, author: 'Fixture Author', text: 'Legacy note text' },
			{
				address: { row: 2, col: 0 },
				author: 'Fixture Person',
				text: 'Threaded root',
				replies: [{ author: 'Fixture Person', text: 'A reply', date: expect.any(String) }],
			},
		]);
	});

	it('reads tables, charts, pictures and unsupported shapes', async () => {
		const sheet = (await loadXlsx(fixture('excel-features.xlsx'))).sheets[1]!;
		expect(sheet.tables[0]).toMatchObject({
			name: 'SalesTable',
			totalsRow: true,
			columns: [
				{ name: 'Region', totalsRowLabel: 'Total' },
				{ name: 'Sales' },
				{ name: 'Cost', totalsRowFunction: 'sum' },
			],
			styleName: 'TableStyleMedium2',
		});
		expect(at(await loadXlsx(fixture('excel-features.xlsx')), 1, 'C6')?.formula).toBe(
			'SUBTOTAL(109,SalesTable[Cost])',
		);
		const [chart, image, shape] = sheet.drawings;
		expect(chart).toMatchObject({
			kind: 'chart',
			chartType: 'column',
			grouping: 'clustered',
			title: 'Sales by region',
			showLegend: true,
		});
		expect(chart?.kind === 'chart' && chart.series[1]).toMatchObject({
			name: 'Cost',
			valuesRef: 'Second!$C$2:$C$5',
			values: [60, 70, 80, 90],
		});
		expect(image).toMatchObject({
			kind: 'image',
			partName: 'xl/media/image1.png',
			contentType: 'image/png',
		});
		expect(shape).toMatchObject({ kind: 'unsupported', description: 'shape' });
		expect(shape?.kind === 'unsupported' && shape.sourceXml).toContain('xdr:sp');
	});
});

describe('loadXlsx with openpyxl and 1904 workbooks', () => {
	it('reads openpyxl styles, formulas with _xlfn prefixes and error literals', async () => {
		const wb = await loadXlsx(fixture('openpyxl-styles.xlsx'));
		expect(at(wb, 0, 'C6')?.formula).toBe('XLOOKUP("Beta",A2:A4,B2:B4)');
		expect(at(wb, 0, 'A7')?.value).toBe('  padded  ');
		expect(styleAt(wb, at(wb, 0, 'A6')?.styleId).alignment).toEqual({
			wrapText: true,
			indent: 1,
			textRotation: 45,
		});
		expect(styleAt(wb, at(wb, 0, 'E2')?.styleId).protection).toEqual({ locked: false });
		expect(styleAt(wb, at(wb, 0, 'E3')?.styleId).fill).toEqual({
			type: 'pattern',
			pattern: 'darkGrid',
			fgColor: { rgb: '00FF0000' },
			bgColor: { rgb: '0000FF00' },
		});
		expect(styleAt(wb, at(wb, 0, 'E4')?.styleId).font).toMatchObject({
			italic: true,
			underline: 'double',
			strike: true,
			vertAlign: 'superscript',
		});
		expect(wb.sheets[0]!.columns.find((c) => c.min === 5)?.hidden).toBe(true);
	});

	it('reads protection, auto filter, page setup and text rules from openpyxl', async () => {
		const sheet = (await loadXlsx(fixture('openpyxl-features.xlsx'))).sheets[0]!;
		expect(sheet.protection).toMatchObject({ sheet: true });
		expect(sheet.protection?.allow).toContain('formatCells');
		expect(sheet.autoFilter?.range).toEqual({ start: { row: 0, col: 0 }, end: { row: 9, col: 1 } });
		expect(sheet.pageSetup).toMatchObject({
			orientation: 'landscape',
			paperSize: 9,
			header: '&CReport',
		});
		expect(
			sheet.conditionalFormats.flatMap((f) => f.rules).find((r) => r.type === 'containsText'),
		).toMatchObject({ text: 'err' });
		expect(sheet.comments.map((c) => c.text)).toEqual(['A note\nsecond line', 'Another']);
	});

	it('flags the 1904 date system', async () => {
		expect((await loadXlsx(fixture('openpyxl-1904.xlsx'))).date1904).toBe(true);
		const excel = await loadXlsx(fixture('excel-1904.xlsx'));
		expect(excel.date1904).toBe(true);
		expect(at(excel, 0, 'A2')?.value).toBe(42521);
	});
});

describe('loadXlsx package handling', () => {
	it('rejects input that is not a zip package', async () => {
		await expect(loadXlsx(new TextEncoder().encode('not a zip'))).rejects.toThrow(
			/Not a valid XLSX package/,
		);
	});

	it('rejects packages without a workbook part', async () => {
		const zip = new JSZip();
		zip.file('hello.txt', 'hi');
		await expect(loadXlsx(await zip.generateAsync({ type: 'uint8array' }))).rejects.toThrow(
			/no workbook part/,
		);
	});

	it('rejects oversized input before unzipping', async () => {
		await expect(loadXlsx(new Uint8Array(51 * 1024 * 1024))).rejects.toThrow(/50 MiB/);
	});

	it('reads inline strings, t="d" dates and case-mismatched part names', async () => {
		const zip = await JSZip.loadAsync(fixture('excel-1904.xlsx'));
		const sheet = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
		zip.file(
			'xl/worksheets/sheet1.xml',
			sheet.replace(
				/<sheetData>[\s\S]*<\/sheetData>/,
				'<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>inline _x005F_x0041_</t></is></c><c r="B1" t="d"><v>1904-01-02T12:00:00</v></c></row></sheetData>',
			),
		);
		const strings = await zip.file('xl/styles.xml')!.async('uint8array');
		zip.remove('xl/styles.xml');
		zip.file('xl/Styles.xml', strings);
		const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		expect(at(wb, 0, 'A1')?.value).toBe('inline _x0041_');
		expect(at(wb, 0, 'B1')?.value).toBe(1.5);
		expect(wb.styles.length).toBeGreaterThan(1);
	});
});
