import { describe, expect, it } from 'vitest';
import { MAX_COL, MAX_ROW } from '../address.js';
import { putCell } from '../cells.js';
import type { Worksheet } from '../model.js';
import { createWorksheet } from '../workbook.js';
import { navigate } from './navigate.js';

/** Column A: rows 2-4 filled, row 7, rows 10-11. Row 0 col B..D filled. */
function sheet(): Worksheet {
	const s = createWorksheet('S', 1);
	for (const r of [2, 3, 4, 7, 10, 11]) putCell(s, r, 0, { value: r });
	for (const c of [1, 2, 3, 6]) putCell(s, 0, c, { value: 'x' });
	return s;
}

const at = (row: number, col: number) => ({ row, col });

describe('navigate arrows', () => {
	const s = sheet();
	s.rowInfo.set(5, { hidden: true });
	s.columns = [{ min: 4, max: 4, hidden: true }];

	it('moves one cell and clamps at the edges', () => {
		expect(navigate(s, at(0, 0), 'up')).toEqual(at(0, 0));
		expect(navigate(s, at(0, 0), 'left')).toEqual(at(0, 0));
		expect(navigate(s, at(0, 0), 'down')).toEqual(at(1, 0));
		expect(navigate(s, at(0, 0), 'right')).toEqual(at(0, 1));
		expect(navigate(s, at(MAX_ROW, MAX_COL), 'down')).toEqual(at(MAX_ROW, MAX_COL));
	});

	it('skips hidden rows and columns', () => {
		expect(navigate(s, at(4, 0), 'down')).toEqual(at(6, 0));
		expect(navigate(s, at(6, 0), 'up')).toEqual(at(4, 0));
		expect(navigate(s, at(0, 3), 'right')).toEqual(at(0, 5));
	});

	it('pages by visible rows', () => {
		expect(navigate(s, at(0, 0), 'pageDown', 10)).toEqual(at(11, 0));
		expect(navigate(s, at(11, 0), 'pageUp', 10)).toEqual(at(0, 0));
		expect(navigate(s, at(3, 0), 'pageUp', 10)).toEqual(at(0, 0));
	});

	it('handles Home, End, Ctrl+Home and Ctrl+End', () => {
		expect(navigate(s, at(0, 3), 'home')).toEqual(at(0, 0));
		expect(navigate(s, at(0, 0), 'end')).toEqual(at(0, 6));
		expect(navigate(s, at(9, 9), 'ctrlHome')).toEqual(at(0, 0));
		expect(navigate(s, at(0, 0), 'ctrlEnd')).toEqual(at(11, 6));
	});

	it('goes to the first cell below frozen panes on Ctrl+Home', () => {
		const f = sheet();
		f.view.freeze = { rows: 1, cols: 2 };
		expect(navigate(f, at(9, 9), 'ctrlHome')).toEqual(at(1, 2));
	});
});

describe('navigate Ctrl+arrows (Excel jump semantics)', () => {
	const s = sheet();

	it('from an empty cell jumps to the next filled cell', () => {
		expect(navigate(s, at(0, 0), 'ctrlDown')).toEqual(at(2, 0));
		expect(navigate(s, at(5, 0), 'ctrlDown')).toEqual(at(7, 0));
	});

	it('inside a block jumps to the block end', () => {
		expect(navigate(s, at(2, 0), 'ctrlDown')).toEqual(at(4, 0));
		expect(navigate(s, at(4, 0), 'ctrlUp')).toEqual(at(2, 0));
		expect(navigate(s, at(10, 0), 'ctrlDown')).toEqual(at(11, 0));
	});

	it('at a block end jumps to the next block', () => {
		expect(navigate(s, at(4, 0), 'ctrlDown')).toEqual(at(7, 0));
		expect(navigate(s, at(7, 0), 'ctrlDown')).toEqual(at(10, 0));
		expect(navigate(s, at(7, 0), 'ctrlUp')).toEqual(at(4, 0));
	});

	it('goes to the sheet edge past the last filled cell', () => {
		expect(navigate(s, at(11, 0), 'ctrlDown')).toEqual(at(MAX_ROW, 0));
		expect(navigate(s, at(2, 0), 'ctrlUp')).toEqual(at(0, 0));
		expect(navigate(s, at(5, 5), 'ctrlRight')).toEqual(at(5, MAX_COL));
	});

	it('works along rows', () => {
		expect(navigate(s, at(0, 0), 'ctrlRight')).toEqual(at(0, 1));
		expect(navigate(s, at(0, 1), 'ctrlRight')).toEqual(at(0, 3));
		expect(navigate(s, at(0, 3), 'ctrlRight')).toEqual(at(0, 6));
		expect(navigate(s, at(0, 6), 'ctrlLeft')).toEqual(at(0, 3));
		expect(navigate(s, at(0, 3), 'ctrlLeft')).toEqual(at(0, 1));
	});

	it('skips hidden cells when jumping', () => {
		const h = sheet();
		h.rowInfo.set(7, { hidden: true });
		expect(navigate(h, at(4, 0), 'ctrlDown')).toEqual(at(10, 0));
		h.rowInfo.set(3, { hidden: true });
		// Rows 2 and 4 are adjacent once row 3 is hidden.
		expect(navigate(h, at(2, 0), 'ctrlDown')).toEqual(at(4, 0));
	});

	it('treats formulas with empty results as filled and empty strings as empty', () => {
		const f = createWorksheet('S', 1);
		putCell(f, 3, 0, { value: '', formula: 'IF(1,"")' });
		putCell(f, 6, 0, { value: '' });
		expect(navigate(f, at(0, 0), 'ctrlDown')).toEqual(at(3, 0));
		expect(navigate(f, at(3, 0), 'ctrlDown')).toEqual(at(MAX_ROW, 0));
	});
});
