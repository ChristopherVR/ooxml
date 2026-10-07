import { describe, expect, it } from 'vitest';
import { calc, calcArray, E } from './test-helpers.js';

// Recorded independently with Excel 16.0 build 20430, 2026-10-07.
describe('lookup omitted arguments and blank references', () => {
	it.each([
		['XMATCH(1,{1},0,)', 1],
		['XMATCH(1,{1},0,A1)', E.VALUE],
		['XMATCH(1,{1},0,IF(TRUE,A1,1))', E.VALUE],
		['XMATCH(1,{1},0,"")', E.VALUE],
		['XMATCH(1,{1},,0)', E.VALUE],
		['MATCH(2,{1;3},)', E.NA],
		['MATCH(2,{1;3},A1)', E.NA],
		['VLOOKUP(2,{1,10;3,30},2,)', E.NA],
		['VLOOKUP(2,{1,10;3,30},2,A1)', E.NA],
		['XLOOKUP(9,{1},{10},)', E.NA],
		['XLOOKUP(9,{1},{10},A1)', 0],
		['XLOOKUP(9,{1},{10},IF(TRUE,A1,1))', 0],
	] as const)('%s matches recorded Excel behavior', (formula, expected) => {
		expect(calc(formula)).toEqual(expected);
	});

	it('retains dynamic-array lifting with an omitted search mode', () => {
		expect(calcArray('XMATCH({1;3},{1;2;3},0,)', {}, 2)).toEqual([[1], [3]]);
	});
});
