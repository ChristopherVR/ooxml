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
	session.setAutoFilter(0, parseRange('A1:A4')!);
	return { workbook, session, sheet: workbook.sheets[0]! };
}

describe('filter row state lifecycle', () => {
	it.each([
		[1, 20],
		[2, 3],
		[3, 3],
		[4, 30],
		[5, 10],
		[6, 6000],
		[7, 10],
		[8, Math.sqrt(200 / 3)],
		[9, 60],
		[10, 100],
		[11, 200 / 3],
	])('SUBTOTAL %s includes rows after deleting the filter range', (code, expected) => {
		const { session, sheet } = setup();
		session.setCellInput(0, 5, 1, `=SUBTOTAL(${code},B2:B4)`);
		session.filterColumn(0, 0, ['keep']);
		session.deleteColumns(0, 0, 1);
		expect(sheet.autoFilter).toBeUndefined();
		expect(getCell(sheet, 5, 0)?.value).toBeCloseTo(expected!);
		expect(sheet.rowInfo.get(2)?.hidden).toBeUndefined();
	});

	it('restores manual hiding through filter deletion, undo, redo and save/reload', async () => {
		const { workbook, session, sheet } = setup();
		session.setHidden(0, 'row', [2, 3], true);
		session.filterColumn(0, 0, ['keep']);
		session.deleteColumns(0, 0, 1);
		expect(sheet.autoFilter).toBeUndefined();
		expect(sheet.rowInfo.get(2)).toEqual({ hidden: true });
		expect(sheet.rowInfo.get(3)).toEqual({ hidden: true });
		expect(getCell(sheet, 5, 0)?.value).toBe(60);
		expect(getCell(sheet, 6, 0)?.value).toBe(10);
		session.undo();
		expect(sheet.autoFilter?.range).toEqual(parseRange('A1:A4'));
		expect(sheet.rowInfo.get(2)?.filteredOut).toBe(true);
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
		session.redo();
		expect(sheet.autoFilter).toBeUndefined();
		expect(getCell(sheet, 5, 0)?.value).toBe(60);
		const loaded = await loadXlsx(await saveXlsx(workbook));
		const reopened = createEditSession(loaded, { autoRowHeight: false });
		reopened.calculateNow({ full: true });
		expect(loaded.sheets[0]?.autoFilter).toBeUndefined();
		expect(loaded.sheets[0]?.rowInfo.get(2)?.hidden).toBe(true);
		expect(getCell(loaded.sheets[0]!, 5, 0)?.value).toBe(60);
		expect(getCell(loaded.sheets[0]!, 6, 0)?.value).toBe(10);
	});

	it.each([false, true])('cleans an unhidden row while preserving metadata: %s', (hasMetadata) => {
		const { session, sheet } = setup();
		session.filterColumn(0, 0, ['keep']);
		if (hasMetadata) session.setRowHeight(0, [1], 22);
		const original = sheet.rowInfo.get(1);
		session.setHidden(0, 'row', [1], true);
		session.setHidden(0, 'row', [1], false);
		expect(sheet.rowInfo.get(1)).toEqual(original);
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
		expect(getCell(sheet, 6, 1)?.value).toBe(40);
		session.undo();
		expect(sheet.rowInfo.get(1)?.hidden).toBe(true);
		session.redo();
		expect(sheet.rowInfo.get(1)).toEqual(original);
	});

	it('keeps an excluded row hidden when manual hiding is cancelled', () => {
		const { session, sheet } = setup();
		session.filterColumn(0, 0, ['keep']);
		session.setHidden(0, 'row', [2], true);
		session.setHidden(0, 'row', [2], false);
		expect(sheet.rowInfo.get(2)).toEqual({ hidden: true, filteredOut: true });
		expect(getCell(sheet, 5, 1)?.value).toBe(40);
		expect(getCell(sheet, 6, 1)?.value).toBe(40);
	});
});
