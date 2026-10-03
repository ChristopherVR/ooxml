// Recalculation cost on large inputs: criteria ranges shared by many formulas, many spills, and
// aggregates over long lists.
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells.js';
import type { Workbook } from '../model.js';
import { book, engine, get } from './test-helpers.js';

function sheetOf(wb: Workbook) {
	const ws = wb.sheets[0];
	if (!ws) throw new Error('no sheet');
	return ws;
}

/** Milliseconds `fn` takes. */
function time(fn: () => void): number {
	const start = performance.now();
	fn();
	return performance.now() - start;
}

describe('criteria functions over shared ranges', () => {
	it('evaluates 2000 SUMIF over the same whole columns in under 500ms', () => {
		const n = 2000;
		const wb = book({});
		const ws = sheetOf(wb);
		for (let r = 0; r < n; r++) {
			putCell(ws, r, 0, { value: r % 10 });
			putCell(ws, r, 1, { value: r });
			putCell(ws, r, 3, { value: null, formula: `SUMIF(A:A,A${r + 1},B:B)` });
		}
		expect(time(() => engine(wb).recalculateAll())).toBeLessThan(500);
		// Rows 0, 10, 20, ...: the sum of 0..1999 stepping by 10.
		expect(get(wb, 'D1')).toBe(199_000);
	});

	it('evaluates 2000 COUNTIF and COUNTIFS over the same columns in under 500ms', () => {
		const n = 2000;
		const wb = book({});
		const ws = sheetOf(wb);
		for (let r = 0; r < n; r++) {
			putCell(ws, r, 0, { value: r % 10 });
			putCell(ws, r, 1, { value: r });
			putCell(ws, r, 3, { value: null, formula: `COUNTIF(A:A,A${r + 1})` });
			putCell(ws, r, 4, { value: null, formula: `COUNTIFS(A:A,A${r + 1},B:B,">=1000")` });
		}
		expect(time(() => engine(wb).recalculateAll())).toBeLessThan(500);
		expect([get(wb, 'D1'), get(wb, 'E1')]).toEqual([200, 100]);
	});
});

describe('spills', () => {
	it('recalculates 8000 spilling formulas and their readers in under 200ms', () => {
		const n = 8000;
		const wb = book({});
		const ws = sheetOf(wb);
		for (let r = 0; r < n; r++) {
			putCell(ws, r, 0, { value: r });
			putCell(ws, r, 1, { value: r });
			putCell(ws, r, 3, { value: null, formula: `A${r + 1}:B${r + 1}*2` });
			putCell(ws, r, 6, { value: null, formula: `E${r + 1}+1` });
		}
		const e = engine(wb);
		e.recalculateAll();
		expect(get(wb, `G${n}`)).toBe((n - 1) * 2 + 1);
		expect(time(() => e.recalculateAll())).toBeLessThan(200);
		expect(get(wb, `G${n}`)).toBe((n - 1) * 2 + 1);
	});
});

describe('aggregates over long lists', () => {
	it('does not spread 200k values into arguments', () => {
		const wb = book({});
		const ws = sheetOf(wb);
		for (let r = 0; r < 200_000; r++) putCell(ws, r, 0, { value: r });
		const formulas = [
			'MAX(A1:A200000)',
			'MIN(A:A)',
			'SUBTOTAL(4,A1:A200000)',
			'SUBTOTAL(105,A:A)',
			'AGGREGATE(4,6,A:A)',
			'MAXIFS(A:A,A:A,">5")',
			'LARGE(A:A,1)',
		];
		formulas.forEach((f, i) => putCell(ws, i, 2, { value: null, formula: f }));
		expect(() => engine(wb).recalculateAll()).not.toThrow();
		expect(formulas.map((_, i) => get(wb, `C${i + 1}`))).toEqual([
			199_999, 0, 199_999, 0, 199_999, 199_999, 199_999,
		]);
	});
});
