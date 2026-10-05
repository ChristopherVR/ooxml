import { describe, expect, it } from 'vitest';
import { parseAddress, parseRange } from '../address.js';
import { getCell } from '../cells.js';
import type { CellValue, Workbook, Worksheet } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { currentRegion } from './filter.js';
import { createEditSession } from './session.js';
import { compareCellValues } from './sort.js';

const A = (ref: string) => {
	const a = parseAddress(ref);
	if (!a) throw new Error(ref);
	return a;
};
const R = (ref: string) => {
	const r = parseRange(ref);
	if (!r) throw new Error(ref);
	return r;
};
const ws = (wb: Workbook): Worksheet => {
	const sheet = wb.sheets[0];
	if (!sheet) throw new Error('sheet');
	return sheet;
};
const column = (wb: Workbook, col: number, from: number, to: number): CellValue[] =>
	Array.from({ length: to - from + 1 }, (_v, i) => getCell(ws(wb), from + i, col)?.value ?? null);
const setup = (rows: CellValue[][]) => {
	const wb = createWorkbook();
	const s = createEditSession(wb, { recalc: false });
	s.setRangeValues(0, A('A1'), rows);
	return { wb, s };
};

describe('compareCellValues', () => {
	it('orders numbers before text before booleans before errors', () => {
		const values: CellValue[] = [{ error: '#N/A' }, true, 'b', 2, 'A', 1];
		values.sort((a, b) => compareCellValues(a, b));
		expect(values).toEqual([1, 2, 'A', 'b', true, { error: '#N/A' }]);
	});
	it('keeps blanks last in both directions', () => {
		const values: CellValue[] = [null, 3, '', 1];
		values.sort((a, b) => compareCellValues(a, b, true));
		expect(values.slice(0, 2)).toEqual([3, 1]);
		expect(values.slice(2).every((v) => v === null || v === '')).toBe(true);
	});
	it('compares text case-insensitively', () => {
		expect(compareCellValues('apple', 'APPLE')).toBe(0);
	});
});

describe('sort', () => {
	it('sorts ascending and descending, moving whole rows', () => {
		const { wb, s } = setup([
			[3, 'c'],
			[1, 'a'],
			[2, 'b'],
		]);
		s.sort(0, R('A1:B3'), [{ col: 0 }], false);
		expect(column(wb, 0, 0, 2)).toEqual([1, 2, 3]);
		expect(column(wb, 1, 0, 2)).toEqual(['a', 'b', 'c']);
		s.sort(0, R('A1:B3'), [{ col: 1, descending: true }], false);
		expect(column(wb, 0, 0, 2)).toEqual([3, 2, 1]);
	});
	it('keeps the header row in place', () => {
		const { wb, s } = setup([['Value'], [2], [1]]);
		s.sort(0, R('A1:A3'), [{ col: 0 }], true);
		expect(column(wb, 0, 0, 2)).toEqual(['Value', 1, 2]);
	});
	it('is stable for equal keys and uses secondary keys', () => {
		const { wb, s } = setup([
			['b', 1],
			['a', 2],
			['b', 0],
			['a', 1],
		]);
		s.sort(0, R('A1:B4'), [{ col: 0 }], false);
		expect(column(wb, 1, 0, 3)).toEqual([2, 1, 1, 0]);
		s.undo();
		s.sort(0, R('A1:B4'), [{ col: 0 }, { col: 1 }], false);
		expect(column(wb, 1, 0, 3)).toEqual([1, 2, 0, 1]);
	});
	it('only moves cells inside the range', () => {
		const { wb, s } = setup([
			[2, 'x'],
			[1, 'y'],
		]);
		s.sort(0, R('A1:A2'), [{ col: 0 }], false);
		expect(column(wb, 0, 0, 1)).toEqual([1, 2]);
		expect(column(wb, 1, 0, 1)).toEqual(['x', 'y']);
	});
	it('translates relative formulas with their rows', () => {
		const { wb, s } = setup([
			[2, null],
			[1, null],
		]);
		s.setCellInput(0, 0, 1, '=A1*10');
		s.setCellInput(0, 1, 1, '=A2*10');
		s.sort(0, R('A1:B2'), [{ col: 0 }], false);
		expect(getCell(ws(wb), 0, 1)?.formula).toBe('A1*10');
		expect(getCell(ws(wb), 1, 1)?.formula).toBe('A2*10');
	});
	it('sorts blanks to the bottom', () => {
		const { wb, s } = setup([[null], [2], [1]]);
		s.setCellValue(0, 0, 0, null);
		s.sort(0, R('A1:A3'), [{ col: 0, descending: true }], false);
		expect(column(wb, 0, 0, 2)).toEqual([2, 1, null]);
	});
	it('undoes a sort', () => {
		const { wb, s } = setup([[3], [1], [2]]);
		s.sort(0, R('A1:A3'), [{ col: 0 }], false);
		s.undo();
		expect(column(wb, 0, 0, 2)).toEqual([3, 1, 2]);
	});
	it('refuses ranges with merged cells', () => {
		const { s } = setup([[3], [1]]);
		s.merge(0, R('A1:B1'), 'merge');
		expect(() => s.sort(0, R('A1:B2'), [{ col: 0 }], false)).toThrow(/merged/);
	});
	it('sorts whole columns up to the used area', () => {
		const { wb, s } = setup([['h'], [5], [4]]);
		s.sort(0, R('A:A'), [{ col: 0 }], true);
		expect(column(wb, 0, 0, 2)).toEqual(['h', 4, 5]);
	});
});

describe('auto-filter', () => {
	const data: CellValue[][] = [
		['Fruit', 'Qty'],
		['apple', 3],
		['pear', 1],
		['apple', 2],
		[null, 5],
	];
	it('expands a single cell to the current region', () => {
		const { wb, s } = setup(data);
		expect(currentRegion(ws(wb), A('B2'))).toEqual(R('A1:B5'));
		s.setAutoFilter(0, R('A1'));
		expect(ws(wb).autoFilter?.range).toEqual(R('A1:B5'));
	});
	it('hides rows that do not match and shows them again', () => {
		const { wb, s } = setup(data);
		s.setAutoFilter(0, R('A1:B5'));
		s.filterColumn(0, 0, ['apple']);
		expect([1, 2, 3, 4].map((r) => !!ws(wb).rowInfo.get(r)?.hidden)).toEqual([
			false,
			true,
			false,
			true,
		]);
		s.filterColumn(0, 0, ['pear', '']);
		expect([1, 2, 3, 4].map((r) => !!ws(wb).rowInfo.get(r)?.hidden)).toEqual([
			true,
			false,
			true,
			false,
		]);
		s.filterColumn(0, 0, undefined);
		expect(ws(wb).rowInfo.size).toBe(0);
	});
	it('combines column filters', () => {
		const { wb, s } = setup(data);
		s.setAutoFilter(0, R('A1:B5'));
		s.filterColumn(0, 0, ['apple']);
		s.filterColumn(0, 1, ['2']);
		expect([1, 2, 3, 4].map((r) => !!ws(wb).rowInfo.get(r)?.hidden)).toEqual([
			true,
			true,
			false,
			true,
		]);
	});
	it('removing the filter unhides rows', () => {
		const { wb, s } = setup(data);
		s.setAutoFilter(0, R('A1:B5'));
		s.filterColumn(0, 0, ['pear']);
		s.setAutoFilter(0, undefined);
		expect(ws(wb).autoFilter).toBeUndefined();
		expect(ws(wb).rowInfo.size).toBe(0);
	});
	it('refuses columns outside the filter', () => {
		const { s } = setup(data);
		expect(() => s.filterColumn(0, 0, ['x'])).toThrow();
	});
	it('sorts by a filter column keeping the header', () => {
		const { wb, s } = setup(data);
		s.setAutoFilter(0, R('A1:B5'));
		s.sortByColumn(0, 1, true);
		expect(column(wb, 1, 0, 4)).toEqual(['Qty', 5, 3, 2, 1]);
	});
	it('sorts a table by column', () => {
		const { wb, s } = setup(data.slice(0, 4));
		s.createTable(0, R('A1:B4'), true);
		s.sortByColumn(0, 1);
		expect(column(wb, 1, 0, 3)).toEqual(['Qty', 1, 2, 3]);
	});
});
