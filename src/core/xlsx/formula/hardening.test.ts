import { describe, expect, it } from 'vitest';
import { hasNestedQuantifier } from './functions/text-regex';
import { compileWildcard } from './functions/wildcard';
import { book, calc, E, engine, get } from './test-helpers';
import { Matrix, resetArrayBudget } from './values';

describe('defined name work limits', () => {
	it('evaluates a 31-level doubling chain quickly', () => {
		const wb = book({ A1: '=Dbl_31' });
		wb.definedNames.push({ name: 'Dbl_0', formula: '1' });
		for (let i = 1; i <= 31; i++)
			wb.definedNames.push({ name: `Dbl_${i}`, formula: `Dbl_${i - 1}+Dbl_${i - 1}` });
		const started = Date.now();
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(2 ** 31);
		expect(Date.now() - started).toBeLessThan(1000);
	});

	it('returns an error for a self-referencing name', () => {
		const wb = book({ A1: '=Loop' });
		wb.definedNames.push({ name: 'Loop', formula: 'Loop+Loop' });
		const started = Date.now();
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.NAME);
		expect(Date.now() - started).toBeLessThan(1000);
	});

	it('bounds a volatile doubling chain with the work budget', () => {
		const wb = book({ A1: '=Vol_30' });
		wb.definedNames.push({ name: 'Vol_0', formula: 'RAND()' });
		for (let i = 1; i <= 30; i++)
			wb.definedNames.push({ name: `Vol_${i}`, formula: `Vol_${i - 1}+Vol_${i - 1}` });
		const started = Date.now();
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.NAME);
		expect(Date.now() - started).toBeLessThan(2000);
	});
});

describe('wildcards without backtracking', () => {
	it('stays fast on a pathological pattern', () => {
		const text = 'a'.repeat(5000);
		const started = Date.now();
		expect(compileWildcard('*a*a*a*a*a*a*a*a*b').test(text)).toBe(false);
		expect(compileWildcard('*a*a*a*a*a*a*a*a*b').search(text)).toBe(-1);
		expect(Date.now() - started).toBeLessThan(1000);
		expect(calc('SEARCH("*a*a*a*a*a*a*a*a*b",A1)', { A1: text })).toEqual(E.VALUE);
	});

	it('matches *, ? and ~ escapes, ignoring case', () => {
		expect(compileWildcard('a*c').test('ABBBC')).toBe(true);
		expect(compileWildcard('a?c').test('abbc')).toBe(false);
		expect(compileWildcard('~*a').test('*A')).toBe(true);
		expect(compileWildcard('~*a').test('xa')).toBe(false);
		expect(compileWildcard('*').test('')).toBe(true);
		expect(compileWildcard('a*b*c').test('axxbxxc')).toBe(true);
		expect(compileWildcard('a*bc').test('abc')).toBe(true);
		expect(compileWildcard('ab*bc').test('abc')).toBe(false);
		expect(compileWildcard('b?d').search('abcbxd', 0)).toBe(3);
		expect(compileWildcard('*d').search('abcd', 2)).toBe(2);
	});
});

describe('regex function limits', () => {
	it('rejects nested quantifiers and long patterns', () => {
		expect(hasNestedQuantifier('(a+)+')).toBe(true);
		expect(hasNestedQuantifier('(a*)*')).toBe(true);
		expect(hasNestedQuantifier('(a|aa)+')).toBe(true);
		expect(hasNestedQuantifier('((a+)b)*')).toBe(true);
		expect(hasNestedQuantifier('(ab)+c*')).toBe(false);
		expect(hasNestedQuantifier('[(a+)+]')).toBe(false);
		expect(calc('REGEXTEST(A1,"(a+)+$")', { A1: `${'a'.repeat(40)}!` })).toEqual(E.VALUE);
		expect(calc('REGEXTEST("abc",REPT("a",1001))')).toEqual(E.VALUE);
		expect(calc('REGEXTEST("abc","b+")')).toBe(true);
	});
});

describe('total array budget', () => {
	it('fails once a recalculation has built too many cells', () => {
		resetArrayBudget();
		try {
			for (let i = 0; i < 4; i++) Matrix.build(4000, 3000, () => 0);
			expect(() => Matrix.build(4000, 3000, () => 0)).toThrow();
		} finally {
			resetArrayBudget();
		}
	});
});
