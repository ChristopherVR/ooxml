import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import { isSpilledCell } from '../formula/spill.js';
import type { Workbook } from '../model.js';
import { createWorkbook } from '../workbook.js';
import { createEditSession } from './session.js';

const val = (wb: Workbook, row: number, col: number, sheet = 0) =>
	getCell(wb.sheets[sheet]!, row, col)?.value ?? null;
const column = (wb: Workbook, col: number, rows: number) =>
	Array.from({ length: rows }, (_v, r) => val(wb, r, col));
const range = (r1: number, c1: number, r2 = r1, c2 = c1) => ({
	start: { row: r1, col: c1 },
	end: { row: r2, col: c2 },
});

function spilling() {
	const wb = createWorkbook();
	const s = createEditSession(wb);
	s.setCellInput(0, 0, 0, '=SEQUENCE(3)');
	return { wb, s };
}

describe('editing dynamic arrays', () => {
	it('typing into a spill range blocks it with #SPILL! and undo restores the spill', () => {
		const { wb, s } = spilling();
		expect(column(wb, 0, 3)).toEqual([1, 2, 3]);
		s.setCellInput(0, 1, 0, 'x');
		expect(val(wb, 0, 0)).toEqual({ error: '#SPILL!' });
		expect(isSpilledCell(getCell(wb.sheets[0]!, 1, 0))).toBe(false);
		s.undo();
		expect(column(wb, 0, 3)).toEqual([1, 2, 3]);
		s.redo();
		expect(val(wb, 0, 0)).toEqual({ error: '#SPILL!' });
		s.clearRange(0, range(1, 0), 'contents');
		expect(column(wb, 0, 3)).toEqual([1, 2, 3]);
	});

	it('replace skips spilled results', () => {
		const { wb, s } = spilling();
		s.setCellValue(0, 0, 3, 2);
		expect(s.replaceAll({ text: '2', wholeCell: true }, '9')).toBe(1);
		expect(val(wb, 0, 3)).toBe(9);
		expect(column(wb, 0, 3)).toEqual([1, 2, 3]);
		const match = { sheet: 0, row: 1, col: 0, text: '2' };
		expect(s.replaceOne({ text: '2' }, '7', match)).toBe(false);
	});

	it('copying the anchor pastes a spilling formula; copying spilled cells alone pastes values', () => {
		const { wb, s } = spilling();
		s.paste(0, { row: 0, col: 2 }, s.copy(0, range(0, 0, 2, 0)));
		expect(getCell(wb.sheets[0]!, 0, 2)?.formula).toBe('SEQUENCE(3)');
		expect(column(wb, 2, 3)).toEqual([1, 2, 3]);
		expect(isSpilledCell(getCell(wb.sheets[0]!, 1, 2))).toBe(true);
		s.paste(0, { row: 0, col: 4 }, s.copy(0, range(1, 0, 2, 0)));
		expect(getCell(wb.sheets[0]!, 0, 4)).toEqual({ value: 2 });
		expect(s.copy(0, range(0, 0, 2, 0)).tsv).toBe('1\r\n2\r\n3\r\n');
	});

	it('fill copies spilled values as constants', () => {
		const { wb, s } = spilling();
		s.fill(0, range(1, 0), range(1, 0, 1, 2));
		expect(getCell(wb.sheets[0]!, 1, 1)).toMatchObject({ value: 2 });
		expect(isSpilledCell(getCell(wb.sheets[0]!, 1, 1))).toBe(false);
	});

	it('sort refuses a range holding part of a dynamic array', () => {
		const { s } = spilling();
		expect(() => s.sort(0, range(0, 0, 2, 0), [{ col: 0 }], false)).toThrow(/dynamic array/);
	});

	it('structural edits move the anchor and respill', () => {
		const { wb, s } = spilling();
		s.insertRows(0, 0, 2);
		expect(column(wb, 0, 5)).toEqual([null, null, 1, 2, 3]);
		s.deleteRows(0, 0, 1);
		expect(column(wb, 0, 4)).toEqual([null, 1, 2, 3]);
		s.undo();
		s.undo();
		expect(column(wb, 0, 3)).toEqual([1, 2, 3]);
	});

	it('formulas typed in a session are dynamic; loaded legacy ones intersect', () => {
		const wb = createWorkbook();
		const s = createEditSession(wb);
		s.setRangeValues(0, { row: 0, col: 0 }, [[1], [2], [3]]);
		s.setCellInput(0, 1, 1, '=A1:A3');
		expect(column(wb, 1, 4)).toEqual([null, 1, 2, 3]);
		const cell = getCell(wb.sheets[0]!, 1, 1)!;
		cell.legacyFormula = true;
		s.calc.invalidate();
		s.calc.recalculateAll();
		expect(column(wb, 1, 4)).toEqual([null, 2, null, null]);
		s.setCellInput(0, 1, 1, '=A1:A3');
		expect(getCell(wb.sheets[0]!, 1, 1)?.legacyFormula).toBeUndefined();
		expect(column(wb, 1, 4)).toEqual([null, 1, 2, 3]);
	});
});
