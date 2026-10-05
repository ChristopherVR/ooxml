import { describe, expect, it } from 'vitest';
import { parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { loadXlsx } from '../read/load.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from '../write/save.js';
import { createEditSession } from './session.js';

const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const setup = () => {
	const wb = createWorkbook({ sheets: ['Sheet1'] });
	return { wb, s: createEditSession(wb) };
};
const column = (wb: Workbook, col: number, rows: number) =>
	Array.from({ length: rows }, (_v, r) => getCell(wb.sheets[0]!, r, col)?.value ?? null);

describe('outline groups', () => {
	it('groups and ungroups rows', () => {
		const { wb, s } = setup();
		s.groupRows(0, 1, 3);
		s.groupRows(0, 2, 2);
		const info = wb.sheets[0]!.rowInfo;
		expect([1, 2, 3].map((r) => info.get(r)?.outlineLevel)).toEqual([1, 2, 1]);
		expect(s.undoLabel()).toBe('Group rows');
		s.ungroupRows(0, 1, 3);
		expect(info.has(1)).toBe(false);
		expect(info.get(2)?.outlineLevel).toBe(1);
		for (let i = 0; i < 9; i++) s.groupRows(0, 5, 5);
		expect(info.get(5)?.outlineLevel).toBe(7);
	});
	it('groups columns, splitting spans, and round-trips the levels', async () => {
		const { wb, s } = setup();
		s.setColumnWidth(0, [0, 1, 2, 3], 20);
		s.groupColumns(0, 1, 2);
		const levelAt = (c: number) =>
			wb.sheets[0]!.columns.find((x) => c >= x.min && c <= x.max)?.outlineLevel;
		expect([0, 1, 2, 3].map(levelAt)).toEqual([undefined, 1, 1, undefined]);
		expect(wb.sheets[0]!.columns.every((c) => c.width === 20)).toBe(true);
		const back = await loadXlsx(await saveXlsx(wb));
		const backCols = back.sheets[0]!.columns;
		expect(backCols.find((c) => c.min <= 1 && c.max >= 1)?.outlineLevel).toBe(1);
		s.ungroupColumns(0, 1, 2);
		expect([1, 2].map(levelAt)).toEqual([undefined, undefined]);
		expect(s.undoLabel()).toBe('Ungroup columns');
	});
	it('sets the standard column width', async () => {
		const { wb, s } = setup();
		s.setDefaultColumnWidth(0, 12);
		expect(wb.sheets[0]!.defaultColWidth).toBe(12);
		expect(() => s.setDefaultColumnWidth(0, 300)).toThrow(RangeError);
		const back = await loadXlsx(await saveXlsx(wb));
		expect(back.sheets[0]!.defaultColWidth).toBe(12);
		s.undo();
		expect(wb.sheets[0]!.defaultColWidth).toBeUndefined();
	});
});

describe('removeDuplicates', () => {
	it('keeps first occurrences (case-insensitive) and moves survivors up', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, { row: 0, col: 0 }, [
			['Name', 'N'],
			['a', 1],
			['B', 2],
			['A', 3],
			['b', 2],
			['c', 5],
		]);
		s.setCellValue(0, 0, 3, 'outside');
		s.setCellValue(0, 4, 3, 'stays');
		const result = s.removeDuplicates(0, R('A1:B6'), [0], true);
		expect(result).toEqual({ removed: 2, remaining: 3 });
		expect(column(wb, 0, 6)).toEqual(['Name', 'a', 'B', 'c', null, null]);
		expect(column(wb, 1, 6)).toEqual(['N', 1, 2, 5, null, null]);
		expect(getCell(wb.sheets[0]!, 4, 3)?.value).toBe('stays');
		expect(s.undoLabel()).toBe('Remove duplicates');
		s.undo();
		expect(column(wb, 0, 6)).toEqual(['Name', 'a', 'B', 'A', 'b', 'c']);
	});
	it('compares every column when none are given and reports nothing to do', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, { row: 0, col: 0 }, [
			['x', 1],
			['x', 2],
			['x', 1],
		]);
		expect(s.removeDuplicates(0, R('A1:B3'), [], false)).toEqual({ removed: 1, remaining: 2 });
		expect(column(wb, 1, 3)).toEqual([1, 2, null]);
		const before = s.undoLabel();
		expect(s.removeDuplicates(0, R('A1:B3'), [], false)).toEqual({ removed: 0, remaining: 2 });
		expect(s.undoLabel()).toBe(before);
	});
	it('adjusts formulas of moved rows', () => {
		const { wb, s } = setup();
		s.setRangeValues(0, { row: 0, col: 0 }, [[1], [1], [2]]);
		s.setCellInput(0, 2, 1, '=A3*10');
		s.removeDuplicates(0, R('A1:B3'), [0], false);
		expect(getCell(wb.sheets[0]!, 1, 1)?.formula).toBe('A2*10');
		expect(getCell(wb.sheets[0]!, 1, 1)?.value).toBe(20);
	});
});
