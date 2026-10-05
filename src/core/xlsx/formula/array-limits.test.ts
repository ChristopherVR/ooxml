// Array size limits and whole-column arrays.
import { describe, expect, it } from 'vitest';
import { book, calc, E, engine, get, grid, set } from './test-helpers.js';

describe('array size limits', () => {
	it.each([
		['ROWS(SEQUENCE(1048577))', E.VALUE],
		['COLUMNS(SEQUENCE(1,1048577))', E.VALUE],
		['ROWS(RANDARRAY(1048577))', E.VALUE],
		['ROWS(RANDARRAY(1000000,1000000))', E.NUM],
		['ROWS(MUNIT(100000))', E.NUM],
		['ROWS(MAKEARRAY(100000,100000,LAMBDA(r,c,1)))', E.NUM],
		['ROWS(EXPAND(1,10000000,10000000))', E.NUM],
		['ROWS(EXPAND(1,1048577))', E.NUM],
		['ROWS(SEQUENCE(1048576,16385))', E.NUM],
	])('%s is an error, not an allocation', (formula, expected) => {
		const start = performance.now();
		expect(calc(formula)).toEqual(expected);
		expect(performance.now() - start).toBeLessThan(1000);
	});

	it('still builds arrays up to the limits', () => {
		expect(calc('ROWS(SEQUENCE(1048576))')).toBe(1_048_576);
		expect(calc('COLUMNS(SEQUENCE(1,20000))')).toBe(20_000);
	});
});

describe('whole-column arrays', () => {
	const data = { A1: 1, A2: 'x', A3: 3, B1: 5 };

	it.each([
		['ROWS(A:A*1)', 1_048_576],
		['COLUMNS(2:2*1)', 16_384],
		['SUMPRODUCT(--(A:A=""))', 1_048_573],
		['SUM(--(A:A=""))', 1_048_573],
		['SUM(IF(A:A="",1,0))', 1_048_573],
		['SUM(ISBLANK(A:A)*1)', 1_048_573],
		['COUNT(A:A+1)', 1_048_575],
		['SUMPRODUCT((A:A="")*(B:B=""))', 1_048_573],
		['COUNTIFS(A:A,"")', 1_048_573],
		['COUNTIFS(A:A,"<>3")', 1_048_575],
		['ROWS(FILTER(A:A,A:A<>""))', 3],
	])('%s', (formula, expected) => {
		expect(calc(formula, data)).toBe(expected);
	});

	it('spills only the used rows of a whole-column result', () => {
		const wb = book({ A1: 1, A2: 2 });
		set(wb, 'C1', '=A:A*10');
		engine(wb).recalculateAll();
		expect(grid(wb, 'C1:C3')).toEqual([[10], [20], [null]]);
		expect(get(wb, 'C2')).toBe(20);
	});
});
