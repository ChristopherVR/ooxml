import { describe, expect, it } from 'vitest';
import { MAX_ROW, type CellRange, createWorkbook } from '../../index';
import { hasOutlinedRow } from './util';

const rows = (first: number, last: number): CellRange => ({
	start: { row: first, col: 0 },
	end: { row: last, col: 0 },
});

describe('hasOutlinedRow', () => {
	it('finds a grouped row inside any of the ranges only', () => {
		const ws = createWorkbook().sheets[0]!;
		ws.rowInfo.set(5, { height: 30 });
		ws.rowInfo.set(10, { outlineLevel: 1 });
		ws.rowInfo.set(20, { outlineLevel: 0 });
		expect(hasOutlinedRow(ws, [rows(0, 9)])).toBe(false);
		expect(hasOutlinedRow(ws, [rows(11, 30)])).toBe(false);
		expect(hasOutlinedRow(ws, [rows(0, 9), rows(10, 10)])).toBe(true);
		expect(hasOutlinedRow(ws, [rows(0, MAX_ROW)])).toBe(true);
	});

	it('answers for a whole column without visiting every row', () => {
		const ws = createWorkbook().sheets[0]!;
		const start = performance.now();
		for (let i = 0; i < 2000; i++) expect(hasOutlinedRow(ws, [rows(0, MAX_ROW)])).toBe(false);
		expect(performance.now() - start).toBeLessThan(process.env.CI ? 150 : 50);
	});
});
