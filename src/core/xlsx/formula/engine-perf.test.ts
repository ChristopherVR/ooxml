// Recalculation cost on large inputs: criteria ranges shared by many formulas, many spills, and
// aggregates over long lists.
import { describe, expect, it } from 'vitest';
import { putCell } from '../cells';
import type { Workbook } from '../model';
import { book, engine, get } from './test-helpers';

/** CI runners share cores and run other test files in parallel, so the budgets get headroom there. */
const SLACK = process.env.CI ? 3 : 1;

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
		expect(time(() => engine(wb).recalculateAll())).toBeLessThan(500 * SLACK);
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
		expect(time(() => engine(wb).recalculateAll())).toBeLessThan(500 * SLACK);
		expect([get(wb, 'D1'), get(wb, 'E1')]).toEqual([200, 100]);
	});
});

describe('exact lookups over shared ranges', () => {
	const n = 20_000;

	it.each([
		['VLOOKUP', (r: number) => `VLOOKUP(D${r},$A$1:$B${n},2,FALSE)`],
		['INDEX/MATCH', (r: number) => `INDEX($B$1:$B${n},MATCH(D${r},$A$1:$A${n},0))`],
		['XLOOKUP', (r: number) => `XLOOKUP(D${r},$A$1:$A${n},$B$1:$B${n})`],
		['reverse XLOOKUP', (r: number) => `XLOOKUP(D${r},$A$1:$A${n},$B$1:$B${n},,0,-1)`],
	])('evaluates 20000 %s in under 1s', (_name, formula) => {
		const wb = book({});
		const ws = sheetOf(wb);
		for (let r = 0; r < n; r++) {
			putCell(ws, r, 0, { value: `Key ${r}` });
			putCell(ws, r, 1, { value: r });
			putCell(ws, r, 3, { value: `key ${n - 1 - r}` });
			putCell(ws, r, 4, { value: null, formula: formula(r + 1) });
		}
		expect(time(() => engine(wb).recalculateAll())).toBeLessThan(1000 * SLACK);
		expect([get(wb, 'E1'), get(wb, `E${n}`)]).toEqual([n - 1, 0]);
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
		expect(time(() => e.recalculateAll())).toBeLessThan(200 * SLACK);
		expect(get(wb, `G${n}`)).toBe((n - 1) * 2 + 1);
	});
});

describe('aggregates over long lists', () => {
	it('evaluates 2000 identical SUMs over 100k rows in under 500ms', () => {
		const wb = book({});
		const ws = sheetOf(wb);
		for (let r = 0; r < 100_000; r++) putCell(ws, r, 0, { value: 1 });
		for (let r = 0; r < 2000; r++)
			putCell(ws, r, 2, { value: null, formula: `SUM($A$1:$A$100000)+${r}` });
		expect(time(() => engine(wb).recalculateAll())).toBeLessThan(500 * SLACK);
		expect(get(wb, 'C2000')).toBe(101_999);
	});

	it('does not visit every formula for volatiles on each edit', () => {
		const wb = book({});
		const ws = sheetOf(wb);
		for (let r = 0; r < 200_000; r++) putCell(ws, r, 0, { value: null, formula: '1+1' });
		putCell(ws, 0, 5, { value: 1 });
		const e = engine(wb);
		e.recalculateAll();
		const edited = [{ sheet: 0, row: 0, col: 5 }];
		expect(
			time(() => {
				for (let i = 0; i < 2000; i++) e.recalculateFrom(edited);
			}),
		).toBeLessThan(400 * SLACK);
	});

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
