import { expect, test } from 'vitest';
import { loadXlsx } from '../xlsx/read/load';
import { saveXlsx } from '../xlsx/write/save';
import { createXlsxFromSheets } from './outline';
import { readXlsxFilledCells } from './xlsx';

/** A workbook with canary text in a very hidden sheet, a hidden sheet, a hidden row and column. */
async function workbook() {
	const book = await loadXlsx(
		await createXlsxFromSheets([
			{
				name: 'Public',
				rows: [
					['name', 'pay', 'canary-column'],
					['canary-row', 1, 2],
					['Ada', 3, 4],
				],
			},
			{ name: 'Payroll', rows: [['canary-very-hidden', 90000]] },
			{ name: 'Old', rows: [['canary-hidden']] },
		]),
	);
	const [visible, payroll, old] = book.sheets;
	payroll!.state = 'veryHidden';
	old!.state = 'hidden';
	visible!.rowInfo.set(1, { hidden: true });
	visible!.columns.push({ min: 2, max: 2, hidden: true });
	return saveXlsx(book);
}

test('omitting hidden content leaves out hidden sheets, rows and columns and counts them', async () => {
	const read = await readXlsxFilledCells(await workbook(), { hidden: 'omit' });
	const text = JSON.stringify(read);
	expect(text).not.toMatch(/canary|Payroll|Old/);
	expect(read.sheets.map((s) => s.name)).toEqual(['Public']);
	expect(read.sheets[0]!.cells.map((c) => c.value)).toEqual(['name', 'pay', 'Ada', 3]);
	expect(read.hidden).toEqual({ sheets: 2, cells: 5 });
});

test('a hidden sheet asked for by index is left out too', async () => {
	const read = await readXlsxFilledCells(await workbook(), { hidden: 'omit', sheetIndex: 1 });
	expect(read.sheets).toEqual([]);
	expect(read.hidden).toEqual({ sheets: 1, cells: 0 });
});

test('every cell is read by default', async () => {
	const read = await readXlsxFilledCells(await workbook());
	expect(JSON.stringify(read)).toMatch(/canary-very-hidden/);
	expect(read.hidden).toBeUndefined();
});
