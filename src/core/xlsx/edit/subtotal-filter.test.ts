import { describe, expect, it } from 'vitest';
import { parseRange } from '../address';
import { getCell } from '../cells';
import { loadXlsx } from '../read';
import { createWorkbook } from '../workbook';
import { saveXlsx } from '../write';
import { createEditSession } from './session';

function setup() {
	const workbook = createWorkbook();
	const session = createEditSession(workbook, { autoRowHeight: false });
	session.setRangeValues(0, { row: 0, col: 0 }, [
		['Kind', 'Amount'],
		['keep', 10],
		['drop', 20],
		['keep', 30],
	]);
	session.setCellInput(0, 5, 1, '=SUBTOTAL(9,B2:B4)');
	session.setCellInput(0, 6, 1, '=SUBTOTAL(109,B2:B4)');
	return { workbook, session, sheet: workbook.sheets[0]! };
}

describe('SUBTOTAL with filtered and manually hidden rows', () => {
	it.each([
		[1, 20],
		[2, 2],
		[3, 2],
		[4, 30],
		[5, 10],
		[6, 300],
		[7, Math.sqrt(200)],
		[8, 10],
		[9, 40],
		[10, 200],
		[11, 100],
	])('function number %s excludes filtered rows and includes manual hiding', (code, expected) => {
		const { session, sheet } = setup();
		session.setCellInput(0, 5, 1, `=SUBTOTAL(${code},B2:B4)`);
		session.setAutoFilter(0, parseRange('A1:B4')!);
		session.filterColumn(0, 0, ['keep']);
		session.setHidden(0, 'row', [1], true);
		expect(getCell(sheet, 5, 1)?.value).toBeCloseTo(expected!);
	});

	it('preserves manual hiding when filters are applied and cleared', () => {
		const { session, sheet } = setup();
		session.setHidden(0, 'row', [2], true);
		session.setAutoFilter(0, parseRange('A1:B4')!);
		session.filterColumn(0, 0, ['keep']);
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
		session.setAutoFilter(0, undefined);
		expect(sheet.rowInfo.get(2)?.hidden).toBe(true);
		expect(getCell(sheet, 5, 1)?.value).toBe(60);
		expect(getCell(sheet, 6, 1)?.value).toBe(40);
	});

	it('excludes filtered rows for both function numbers, including after undo and redo', () => {
		const { session, sheet } = setup();
		session.setAutoFilter(0, parseRange('A1:B4')!);
		session.filterColumn(0, 0, ['keep']);
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
		expect(getCell(sheet, 6, 1)?.value).toBe(40);
		session.undo();
		expect(getCell(sheet, 5, 1)?.value).toBe(60);
		session.redo();
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
	});

	it('includes manually hidden rows only for function numbers 1-11', () => {
		const { session, sheet } = setup();
		session.setAutoFilter(0, parseRange('A1:B4')!);
		session.filterColumn(0, 0, ['keep']);
		session.setHidden(0, 'row', [1], true);
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
		expect(getCell(sheet, 6, 1)?.value).toBe(30);
	});

	it('keeps a filtered row excluded until the filter is reapplied', () => {
		const { session, sheet } = setup();
		session.setAutoFilter(0, parseRange('A1:B4')!);
		session.filterColumn(0, 0, ['keep']);
		session.setCellInput(0, 2, 0, 'keep');
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
		session.filterColumn(0, 0, ['keep']);
		expect(getCell(sheet, 5, 1)?.value).toBe(60);
	});

	it('does not classify a manually hidden row as filtered just because its value changes', () => {
		const { session, sheet } = setup();
		session.setAutoFilter(0, parseRange('A1:B4')!);
		session.filterColumn(0, 0, ['keep']);
		session.setHidden(0, 'row', [1], true);
		session.setCellInput(0, 1, 0, 'drop');
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
		session.filterColumn(0, 0, ['keep']);
		expect(getCell(sheet, 5, 1)?.value).toBe(30);
	});

	it('retains the distinction after saving and reopening', async () => {
		const { workbook, session } = setup();
		session.setAutoFilter(0, parseRange('A1:B4')!);
		session.filterColumn(0, 0, ['keep']);
		session.setHidden(0, 'row', [1], true);
		const loaded = await loadXlsx(await saveXlsx(workbook));
		createEditSession(loaded, { autoRowHeight: false });
		expect(getCell(loaded.sheets[0]!, 5, 1)?.value).toBe(40);
		expect(getCell(loaded.sheets[0]!, 6, 1)?.value).toBe(30);
	});
});
