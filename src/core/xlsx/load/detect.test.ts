import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import { decodeText, detectWorkbookFormat, loadWorkbook, saveWorkbook } from './detect.js';
import { LegacyXlsError } from './legacy-xls.js';

const fixture = async (path: string): Promise<Uint8Array> =>
	new Uint8Array(await readFile(new URL(`../__fixtures__/${path}`, import.meta.url)));

async function zipWith(names: string[]): Promise<Uint8Array> {
	const zip = new JSZip();
	for (const name of names) zip.file(name, 'x');
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('detectWorkbookFormat', () => {
	it('sniffs compound files, zips and text', async () => {
		expect(detectWorkbookFormat(await fixture('xls/workbook-styles.xls'))).toBe('xls');
		expect(detectWorkbookFormat(await zipWith(['xl/workbook.xml']))).toBe('xlsx');
		expect(detectWorkbookFormat(new TextEncoder().encode('a,b\n1,2\n'))).toBe('csv');
		expect(detectWorkbookFormat(new TextEncoder().encode('a,b').buffer)).toBe('csv');
	});

	it('recognises macro-enabled packages by VBA part or file name', async () => {
		const plain = await zipWith(['xl/workbook.xml']);
		expect(detectWorkbookFormat(await zipWith(['xl/workbook.xml', 'xl/vbaProject.bin']))).toBe(
			'xlsm',
		);
		expect(detectWorkbookFormat(plain, 'Budget.XLSM')).toBe('xlsm');
		// The name never overrides the content.
		expect(detectWorkbookFormat(new TextEncoder().encode('a'), 'book.xlsx')).toBe('csv');
	});

	it('rejects binary data that is not a workbook', () => {
		expect(() => detectWorkbookFormat(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0, 1]))).toThrow(
			/Unsupported file/,
		);
	});
});

describe('decodeText', () => {
	it('handles UTF-8 with and without BOM and UTF-16 with BOM', () => {
		expect(decodeText(new Uint8Array([0xef, 0xbb, 0xbf, 0x61, 0xc3, 0xa9]))).toBe('aé');
		expect(decodeText(new Uint8Array([0xff, 0xfe, 0x61, 0x00, 0x2c, 0x00]))).toBe('a,');
		expect(decodeText(new Uint8Array([0xfe, 0xff, 0x00, 0x61]))).toBe('a');
	});
});

describe('loadWorkbook', () => {
	it('dispatches .xls to the legacy reader', async () => {
		const workbook = await loadWorkbook(await fixture('xls/workbook-features.xls'));
		expect(workbook.format).toBe('xls');
		expect(workbook.sheets).toHaveLength(4);
		await expect(loadWorkbook(await fixture('xls/workbook-encrypted.xls'))).rejects.toThrow(
			LegacyXlsError,
		);
	});

	it('dispatches .xlsx to the package reader', async () => {
		const workbook = await loadWorkbook(await fixture('excel-features.xlsx'));
		expect(workbook.format).toBe('xlsx');
		expect(workbook.sheets.length).toBeGreaterThan(0);
	});

	it('loads CSV with a sheet named after the file and an explicit delimiter', async () => {
		const bytes = new TextEncoder().encode('﻿a|b\n1|2\n');
		const workbook = await loadWorkbook(bytes, { fileName: 'export.csv', delimiter: '|' });
		expect(workbook.format).toBe('csv');
		expect(workbook.sheets[0]?.name).toBe('export');
		expect(getCell(workbook.sheets[0]!, 1, 1)?.value).toBe(2);
	});
});

describe('saveWorkbook', () => {
	it('writes CSV as UTF-8 with a byte-order mark', async () => {
		const workbook = await loadWorkbook(new TextEncoder().encode('x,y\n1,é\n'));
		const bytes = await saveWorkbook(workbook, 'csv');
		expect([...bytes.subarray(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
		expect(new TextDecoder().decode(bytes.subarray(3))).toBe('x,y\r\n1,é\r\n');
	});

	it('saves an .xls workbook as .xlsx that loads back with its values and formulas', async () => {
		const workbook = await loadWorkbook(await fixture('xls/workbook-features.xls'));
		const saved = await saveWorkbook(workbook, 'xlsx');
		expect(detectWorkbookFormat(saved)).toBe('xlsx');
		const reloaded = await loadWorkbook(saved);
		expect(reloaded.sheets.map((sheet) => sheet.name)).toEqual(
			workbook.sheets.map((sheet) => sheet.name),
		);
		const formulas = reloaded.sheets[1]!;
		expect(getCell(formulas, 0, 0)).toMatchObject({ formula: 'SUM(Data!B2:B4)', value: 5.5 });
		expect(getCell(reloaded.sheets[0]!, 7, 0)?.value).toBe('0123456789'.repeat(1000));
	});
});
