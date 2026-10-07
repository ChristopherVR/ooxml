import { describe, expect, it } from 'vitest';
import { calc } from './test-helpers.js';
import { xsearch } from './functions/lookup-core.js';

// Recorded independently with Excel 16.0 build 20430, 2026-10-07.
describe('modern binary lookup search', () => {
	it.each([
		['XMATCH(1,{1;1;1;1;1},0,2)', 1],
		['XMATCH(1,{1;1;1;1;1},0,-2)', 5],
		['XMATCH(1,{1;1;3;3},0,2)', 1],
		['XMATCH(3,{1;1;3;3},0,2)', 3],
		['XMATCH(3,{1;1;3;3},-1,2)', 3],
		['XMATCH(3,{1;1;3;3},1,2)', 3],
		['XMATCH(1,{3;3;1;1},0,-2)', 4],
		['XMATCH(3,{3;3;1;1},-1,-2)', 2],
		['XLOOKUP(3,{1;1;3;3},{10;20;30;40},"missing",0,2)', 30],
		['XMATCH(2,{1;3;"a";FALSE},1,2)', 2],
		['XMATCH("",{1;3;"a";FALSE},-1,2)', 2],
		['XMATCH("",{1;3;"a";FALSE},1,2)', 3],
		['XMATCH(2,{FALSE;"a";3;1},-1,-2)', 4],
	] as const)('%s matches recorded Excel behavior', (formula, expected) => {
		expect(calc(formula)).toBe(expected);
	});

	it.each([
		['XMATCH(2,A1:A6,1,2)', 2],
		['XMATCH("",A1:A6,-1,2)', 2],
		['XMATCH("",A1:A6,1,2)', 3],
		['XMATCH(TRUE,A1:A6,1,2)', 5],
		['XMATCH(A6,A1:A6,0,2)', 5],
	] as const)('%s includes sorted trailing blanks', (formula, expected) => {
		expect(calc(formula, { A1: 1, A2: 3, A3: 'a', A4: false })).toBe(expected);
	});

	it('searches a whole-column blank tail without reading its cells', () => {
		let reads = 0;
		const vector = {
			length: 1,
			logicalLength: 1_048_576,
			get: (i: number) => {
				if (i !== 0) throw new Error('read past stored prefix');
				reads++;
				return 1;
			},
		};
		expect(xsearch(vector, null, 0, 2)).toBe(1);
		expect(xsearch(vector, 4, 1, 2)).toBe(1);
		expect(reads).toBeLessThanOrEqual(4);
	});
});
