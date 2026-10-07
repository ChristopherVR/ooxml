import { describe, expect, it } from 'vitest';
import { calc, E } from './test-helpers.js';
import { findNearest } from './functions/lookup-core.js';

// Recorded independently with Excel 16.0 build 20430, 2026-10-07.
describe('modern approximate lookup ordering', () => {
	it.each([
		['XMATCH("",{1;3},-1)', 2],
		['XMATCH("",{1;3},1)', E.NA],
		['XMATCH("b",{1;"a";"c";TRUE},-1)', 2],
		['XMATCH("b",{1;"a";"c";TRUE},1)', 3],
		['XMATCH(2,{"a";"b"},-1)', E.NA],
		['XMATCH(2,{"a";"b"},1)', 1],
		['XMATCH(TRUE,{1;"a";FALSE},-1)', 3],
		['XMATCH(TRUE,{1;"a";FALSE},1)', E.NA],
		['XMATCH(0,{FALSE;TRUE},1)', 1],
		['XMATCH(FALSE,{0;1},-1)', 2],
		['XMATCH(2,{1;1;3;3},-1,-1)', 2],
		['XMATCH("b",{#N/A;1},-1)', 2],
		['XMATCH("",{3;1},-1,-1)', 1],
		['XMATCH(2,{"c";"a"},1,-1)', 2],
		['XMATCH(A2,{"";TRUE},-1)', 2],
		['XMATCH(A2,{"";TRUE},1)', E.NA],
	] as const)('%s matches recorded Excel behavior', (formula, expected) => {
		expect(calc(formula)).toEqual(expected);
	});

	it.each([
		['XMATCH(A2,A1:A5,-1)', 2],
		['XMATCH(A2,A1:A5,1)', 2],
		['XMATCH(A2,A1:A5,-1,-1)', 5],
		['XMATCH(A2,A1:A5,1,-1)', 5],
		['XMATCH(0,A1:A5,-1)', E.NA],
		['XMATCH(4,A1:A5,1)', 2],
		['XMATCH("",A1:A5,-1)', 3],
		['XMATCH("",A1:A5,1)', 2],
		['XMATCH("",A1:A5,1,-1)', 5],
		['XMATCH(FALSE,A1:A5,-1)', 3],
		['XMATCH(FALSE,A1:A5,1)', 2],
		['XLOOKUP("",A1:A5,{10;20;30;40;50},"missing",-1)', 30],
		['XLOOKUP(A2,A1:A5,{10;20;30;40;50},"missing",-1,-1)', 50],
	] as const)('%s orders blanks after other types', (formula, expected) => {
		expect(calc(formula, { A1: 1, A3: 3 })).toEqual(expected);
	});

	it('considers a whole-column blank tail with bounded reads', () => {
		let reads = 0;
		const vector = {
			length: 1,
			logicalLength: 1_048_576,
			get: () => {
				reads++;
				return 1;
			},
		};
		expect(findNearest(vector, null, -1, true)).toBe(1_048_575);
		expect(reads).toBe(0);
		expect(findNearest(vector, 4, 1, false)).toBe(1);
		expect(reads).toBe(1);
	});
});
