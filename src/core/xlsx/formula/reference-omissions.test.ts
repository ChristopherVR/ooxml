import { describe, expect, it } from 'vitest';
import { calc, calcArray, E } from './test-helpers.js';

// Recorded independently with Excel 16.0 build 20430, 2026-10-07.
describe('reference function omissions', () => {
	it.each([
		['ADDRESS(1,2,)', '$B$1'],
		['ADDRESS(1,2,A2)', E.VALUE],
		['ADDRESS(1,2,1,)', '$B$1'],
		['ADDRESS(1,2,1,A2)', 'R1C2'],
		['ADDRESS(1,2,,A2)', 'R1C2'],
		['ADDRESS(1,2,,,"Sheet 1")', "'Sheet 1'!$B$1"],
		['ADDRESS(1,2,4,FALSE)', 'R[1]C[2]'],
		['INDIRECT("A1",)', E.REF],
		['INDIRECT("A1",A2)', E.REF],
		['INDIRECT("R1C1",A2)', 42],
		['INDIRECT("R1C1",)', 42],
		['ROWS(OFFSET(A1:A2,0,0,,))', 2],
		['ROWS(OFFSET(A1:A2,0,0,B1))', E.REF],
		['ROWS(OFFSET(A1:A2,0,0,IF(TRUE,B1,1)))', E.REF],
		['COLUMNS(OFFSET(A1:B2,0,0,,B3))', E.REF],
		['ROWS(OFFSET(A1:A2,0,0,0))', E.REF],
		['ROWS(OFFSET(A1:A2,0,0,,1))', 2],
		['ROWS(OFFSET(A1:A2,0,0,-1))', 1],
	] as const)('%s matches recorded Excel behavior', (formula, expected) => {
		expect(calc(formula, { A1: 42 })).toEqual(expected);
	});

	it('retains array lifting with omitted ADDRESS defaults', () => {
		expect(calcArray('ADDRESS({1;2},2,,,"Sheet 1")', {}, 2)).toEqual([
			["'Sheet 1'!$B$1"],
			["'Sheet 1'!$B$2"],
		]);
	});

	it('derives omitted OFFSET dimensions from the input reference', () => {
		expect(calc('ROWS(OFFSET(A1:C2,0,0,,))')).toBe(2);
		expect(calc('COLUMNS(OFFSET(A1:C2,0,0,,))')).toBe(3);
	});
});
