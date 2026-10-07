import { describe, expect, it } from 'vitest';
import { calc, E } from './test-helpers';

// Recorded independently with Excel 16.0 build 20430, 2026-10-07.
describe('classic lookup array arguments', () => {
	it.each([
		['MATCH(1,1,0)', E.NA],
		['MATCH("a","a",0)', E.VALUE],
		['MATCH(1,IF(TRUE,1,2),0)', E.NA],
		['VLOOKUP(1,1,1,FALSE)', E.NA],
		['HLOOKUP(1,1,1,FALSE)', E.NA],
		['VLOOKUP(1,1,0,FALSE)', E.NA],
		['MATCH(1,#VALUE!,0)', E.VALUE],
		['VLOOKUP(1,#REF!,1,FALSE)', E.REF],
		['HLOOKUP(1,#DIV/0!,1,FALSE)', E.DIV0],
	] as const)('%s matches recorded Excel behavior', (formula, expected) => {
		expect(calc(formula)).toEqual(expected);
	});

	it('accepts a one-cell reference, array constant and dynamic array', () => {
		expect(calc('MATCH(1,A1,0)', { A1: 1 })).toBe(1);
		expect(calc('MATCH(1,{1},0)')).toBe(1);
		expect(calc('MATCH(1,SEQUENCE(1),0)')).toBe(1);
		expect(calc('VLOOKUP(1,{1},1,FALSE)')).toBe(1);
		expect(calc('HLOOKUP(1,A1,1,FALSE)', { A1: 1 })).toBe(1);
	});

	it('retains the scalar support of XMATCH and LOOKUP', () => {
		expect(calc('XMATCH(1,1)')).toBe(1);
		expect(calc('LOOKUP(1,1)')).toBe(1);
	});
});
