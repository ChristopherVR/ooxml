import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { styleAt } from '../styles.js';
import { LegacyXlsError, loadLegacyXls } from './legacy-xls.js';

// Real Excel 16 saves (FileFormat 56), copied from the ole2 repository where
// test/fixtures/xls/generate-xls-fixtures.ps1 regenerates them.
const fixture = async (name: string): Promise<Uint8Array> =>
	new Uint8Array(await readFile(new URL(`../__fixtures__/xls/${name}`, import.meta.url)));

const cellStyle = (workbook: Workbook, sheet: number, row: number, col: number) => {
	const target = workbook.sheets[sheet];
	return styleAt(workbook, target ? getCell(target, row, col)?.styleId : undefined);
};

describe('loadLegacyXls: values, sheets and formulas', async () => {
	const workbook = await loadLegacyXls(await fixture('workbook-features.xls'));
	const [data, formulas] = workbook.sheets;

	it('maps sheets, visibility and workbook settings', () => {
		expect(workbook.format).toBe('xls');
		expect(workbook.sheets.map((sheet) => [sheet.name, sheet.state, sheet.sheetId])).toEqual([
			['Data', 'visible', 1],
			['My Formulas', 'visible', 2],
			['Hidden', 'hidden', 3],
			['Secret', 'veryHidden', 4],
		]);
		expect(workbook.date1904).toBe(false);
		expect(workbook.activeSheet).toBe(0);
		expect(workbook.warnings[0]).toMatch(/saving writes \.xlsx/);
		expect(workbook.properties.title).toBe('Fixture workbook-features.xls');
	});

	it('keeps typed values including unicode, long strings, booleans and errors', () => {
		if (!data) throw new Error('missing sheet');
		expect(getCell(data, 0, 0)?.value).toBe('Name');
		expect(getCell(data, 4, 0)?.value).toBe('日本語 été Ж');
		expect(getCell(data, 4, 1)?.value).toBe(-1234567.891);
		expect(getCell(data, 1, 3)?.value).toBe(true);
		expect(getCell(data, 1, 4)).toMatchObject({ value: { error: '#DIV/0!' }, formula: '1/0' });
		expect(getCell(data, 7, 0)?.value).toBe('0123456789'.repeat(1000));
		expect(getCell(data, 8, 0)?.value).toBe('ÅΩ中x'.repeat(3000));
	});

	it('keeps date and percentage formats on the cell style', () => {
		expect(getCell(data!, 1, 5)?.value).toBe(45292);
		expect(cellStyle(workbook, 0, 2, 5).numFmt).toBe('h:mm AM/PM');
		expect(cellStyle(workbook, 0, 1, 6).numFmt).toBe('0.00%');
	});

	it('keeps formula text and cached results', () => {
		if (!formulas) throw new Error('missing sheet');
		expect(getCell(formulas, 0, 0)).toMatchObject({ formula: 'SUM(Data!B2:B4)', value: 5.5 });
		expect(getCell(formulas, 1, 0)).toMatchObject({
			formula: 'IF(Data!B2>1,"big","small")',
			value: 'big',
		});
		expect(getCell(formulas, 2, 0)?.formula).toBe('VLOOKUP("Banana",Data!A2:C4,3,FALSE)');
		expect(getCell(formulas, 22, 0)).toMatchObject({ formula: 'IFERROR(1/0,"err")', value: 'err' });
		expect(getCell(formulas, 28, 0)?.formula).toBe(`Secret!A1&" "&'My Formulas'!A2`);
		expect(getCell(formulas, 0, 2)).toMatchObject({
			formula: 'Data!B2:B4*10',
			arrayRange: { start: { row: 0, col: 2 }, end: { row: 2, col: 2 } },
		});
		expect(getCell(formulas, 1, 2)?.formula).toBeUndefined();
		expect(getCell(formulas, 1, 2)?.value).toBe(5);
	});

	it('maps defined names, dropping hidden future-function names', () => {
		expect(workbook.definedNames).toEqual([
			{ name: '_xlnm.Print_Area', formula: 'Data!$A$1:$C$5', localSheet: 0 },
			{ name: 'TaxRate', formula: 'Data!$B$2' },
		]);
	});
});

describe('loadLegacyXls: formatting and layout', async () => {
	const workbook = await loadLegacyXls(await fixture('workbook-styles.xls'));
	const sheet = workbook.sheets[0]!;

	it('interns resolved styles with palette colours as RGB', () => {
		expect(workbook.styles[0]?.font.name).toBe('Aptos Narrow');
		expect(cellStyle(workbook, 0, 0, 0).font).toMatchObject({
			bold: true,
			color: { rgb: 'FFFF0000' },
		});
		expect(cellStyle(workbook, 0, 1, 0).font).toMatchObject({
			name: 'Arial',
			size: 14,
			italic: true,
		});
		expect(cellStyle(workbook, 0, 2, 0).fill).toEqual({
			type: 'pattern',
			pattern: 'solid',
			fgColor: { rgb: 'FFFFFF00' },
		});
		expect(cellStyle(workbook, 0, 1, 1).border).toMatchObject({
			left: { style: 'thin' },
			top: { style: 'medium' },
			bottom: { style: 'double', color: { rgb: 'FF0000FF' } },
			right: { style: 'dashed' },
		});
		expect(cellStyle(workbook, 0, 4, 1).alignment).toEqual({
			horizontal: 'center',
			vertical: 'top',
			wrapText: true,
		});
		expect(cellStyle(workbook, 0, 3, 1).numFmt).toBe('"USD"\\ #,##0.000;[Red]\\-#,##0.000');
		expect(cellStyle(workbook, 0, 6, 0).font.color).toEqual({ theme: 10 });
		const ids = new Set(workbook.styles.map((style) => JSON.stringify(style)));
		expect(ids.size).toBe(workbook.styles.length);
	});

	it('maps merges, columns, rows, view and freeze panes', () => {
		expect(sheet.merges).toEqual([
			{ start: { row: 0, col: 3 }, end: { row: 1, col: 5 } },
			{ start: { row: 3, col: 3 }, end: { row: 3, col: 4 } },
		]);
		expect(sheet.columns.find((col) => col.min === 0)).toMatchObject({ max: 0, customWidth: true });
		expect(sheet.columns.find((col) => col.min === 6)?.hidden).toBe(true);
		expect(sheet.rowInfo.get(1)).toMatchObject({ height: 30, customHeight: true });
		expect(sheet.rowInfo.get(8)?.hidden).toBe(true);
		expect(sheet.view).toMatchObject({
			showGridLines: false,
			zoom: 85,
			freeze: { rows: 2, cols: 1 },
			topLeft: { row: 2, col: 1 },
		});
		expect(workbook.sheets[1]?.tabColor).toEqual({ rgb: 'FFFF0000' });
	});

	it('maps hyperlinks and comments', () => {
		expect(sheet.hyperlinks).toEqual([
			{
				range: { start: { row: 6, col: 0 }, end: { row: 6, col: 0 } },
				target: 'https://example.com/path?q=1',
				tooltip: 'Example tip',
				display: 'Example site',
			},
			{
				range: { start: { row: 7, col: 0 }, end: { row: 7, col: 0 } },
				location: 'Other!B3',
				display: 'Jump to Other',
			},
		]);
		expect(sheet.comments).toEqual([
			{ address: { row: 5, col: 0 }, author: 'Author', text: 'Reviewed by QA' },
		]);
	});
});

describe('loadLegacyXls: modes and errors', () => {
	it('reads the 1904 date system', async () => {
		const workbook = await loadLegacyXls(await fixture('workbook-1904.xls'));
		expect(workbook.date1904).toBe(true);
		expect(getCell(workbook.sheets[0]!, 1, 0)?.value).toBe(43830);
	});

	it('rejects encrypted and unreadable files with LegacyXlsError', async () => {
		const encrypted = loadLegacyXls(await fixture('workbook-encrypted.xls'));
		await expect(encrypted).rejects.toThrow(LegacyXlsError);
		await expect(loadLegacyXls(await fixture('workbook-encrypted.xls'))).rejects.toMatchObject({
			code: 'encrypted',
			message: expect.stringMatching(/password protected/),
		});
		await expect(loadLegacyXls(new Uint8Array(64))).rejects.toMatchObject({ code: 'corrupt' });
	});
});
