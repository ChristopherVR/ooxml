import { expect, it } from 'vitest';
import { createEditSession, createWorkbook, parseRange } from '../index';
import { formatCellsState } from './format-cells-state';
import { stepMatch } from './find-order';

it('reports shared formatting and names the mixed properties', () => {
	const book = createWorkbook();
	const session = createEditSession(book);
	session.applyStyle(0, [parseRange('A1')!], { font: { bold: true, size: 14 }, numFmt: '0.00' });
	session.applyStyle(0, [parseRange('B1')!], { font: { bold: true, size: 12 } });
	const state = formatCellsState(book, 0, [parseRange('A1:B1')!]);
	expect(state.font.value.bold).toBe(true);
	expect(state.font.mixed).toEqual(['size']);
	expect(state.numFmt).toBeNull();
	expect(state.complete).toBe(true);
	expect(formatCellsState(book, 0, [parseRange('A1')!]).numFmt).toBe('0.00');
});
it('bounds whole-sheet selections', () => {
	const book = createWorkbook();
	expect(formatCellsState(book, 0, [parseRange('A1:XFD1048576')!]).complete).toBe(false);
});
it('steps to the next and previous match, wrapping', () => {
	const m = (row: number, col: number) => ({ sheet: 0, row, col, text: '' });
	const all = [m(0, 0), m(2, 0), m(4, 0)];
	expect(stepMatch(all, 0, 2, 0, false)).toEqual(m(4, 0));
	expect(stepMatch(all, 0, 2, 0, false, true)).toEqual(m(0, 0));
	expect(stepMatch(all, 0, 3, 0, false, true)).toEqual(m(2, 0));
	expect(stepMatch(all, 0, 4, 0, false)).toEqual(m(0, 0));
	expect(stepMatch([], 0, 0, 0, false)).toBeUndefined();
});
