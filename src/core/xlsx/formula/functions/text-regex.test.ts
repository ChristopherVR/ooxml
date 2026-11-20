import { describe, expect, it } from 'vitest';
import { calc, calcArray, E } from '../test-helpers';

describe('ENCODEURL', () => {
	it('encodes like Microsoft documents it', () => {
		expect(
			calc('ENCODEURL("http://contoso.sharepoint.com/Finance/Profit and Loss Statement.xlsx")'),
		).toBe('http%3A%2F%2Fcontoso.sharepoint.com%2FFinance%2FProfit%20and%20Loss%20Statement.xlsx');
	});
	it('keeps unreserved characters, encodes the rest as UTF-8', () => {
		expect(calc('ENCODEURL("a-b_c.d~e")')).toBe('a-b_c.d~e');
		expect(calc('ENCODEURL("café (1)*!")')).toBe('caf%C3%A9%20%281%29%2A%21');
	});
});

describe('regular expression functions', () => {
	it('REGEXTEST honours case sensitivity', () => {
		expect(calc('REGEXTEST("Hello","^h")')).toBe(false);
		expect(calc('REGEXTEST("Hello","^h",1)')).toBe(true);
		expect(calc('REGEXTEST("Hello","^h",2)')).toEqual(E.VALUE);
	});

	it('REGEXEXTRACT returns the first match or #N/A', () => {
		expect(calc('REGEXEXTRACT("order 4521 shipped","[0-9]+")')).toBe('4521');
		expect(calc('REGEXEXTRACT("none","[0-9]+")')).toEqual(E.NA);
		expect(calc('REGEXEXTRACT("a1b2","[0-9]",3)')).toEqual(E.VALUE);
	});

	it('spills all matches down rows, retaining numeric text', () => {
		expect(calcArray('REGEXEXTRACT("a01b22c333","[0-9]+",1)', {}, 3, 1)).toEqual([
			['01'],
			['22'],
			['333'],
		]);
		expect(calcArray('REGEXEXTRACT("aAbB","[a-b]",1,1)', {}, 4, 1)).toEqual([
			['a'],
			['A'],
			['b'],
			['B'],
		]);
		expect(calc('REGEXEXTRACT("none","[0-9]+",1)')).toEqual(E.NA);
	});

	it('spills only the first match’s capture groups', () => {
		expect(calcArray('REGEXEXTRACT("John Smith; Jane Doe","(\\w+) (\\w+)",2)', {}, 1, 2)).toEqual([
			['John', 'Smith'],
		]);
		expect(calcArray('REGEXEXTRACT("b","(a)?(b)",2)', {}, 1, 2)).toEqual([['', 'b']]);
		expect(calc('REGEXEXTRACT("none","([0-9]+)",2)')).toEqual(E.NA);
		expect(calc('REGEXEXTRACT("abc","abc",2)')).toEqual(E.VALUE);
	});

	it('advances safely for zero-length and Unicode matches', () => {
		expect(calcArray('REGEXEXTRACT("ab","(?=.)",1)', {}, 2, 1)).toEqual([[''], ['']]);
		expect(calcArray('REGEXEXTRACT("😀a",".",1)', {}, 2, 1)).toEqual([['😀'], ['a']]);
	});

	it('REGEXREPLACE replaces all, the nth or the nth from the end', () => {
		expect(calc('REGEXREPLACE("a1b22c","[0-9]+","#")')).toBe('a#b#c');
		expect(calc('REGEXREPLACE("a1b22c","[0-9]+","#",2)')).toBe('a1b#c');
		expect(calc('REGEXREPLACE("a1b22c","[0-9]+","#",-2)')).toBe('a#b22c');
		expect(calc('REGEXREPLACE("a1b22c","[0-9]+","#",5)')).toBe('a1b22c');
		expect(calc('REGEXREPLACE("John Smith","(\\w+) (\\w+)","$2, $1")')).toBe('Smith, John');
	});

	it('reports an invalid pattern as #VALUE!', () => {
		expect(calc('REGEXTEST("x","(")')).toEqual(E.VALUE);
		expect(calc('REGEXTEST("x","a{2,1}")')).toEqual(E.VALUE);
		expect(calc('REGEXTEST("x","a++")')).toEqual(E.VALUE);
		expect(calc('REGEXTEST("abc",REPT("a",1001))')).toEqual(E.VALUE);
	});

	it('replaces one occurrence with that match’s own captures and context', () => {
		expect(calc('REGEXREPLACE("ab ab","(?<=a)b","[$&$`]",2)')).toBe('ab a[bab a]');
		expect(calc('REGEXREPLACE("x1 y2","(?<l>[a-z])(\\d)","$2$<l>",-1)')).toBe('x1 2y');
	});
});

describe('regular expressions cannot hang a recalculation', () => {
	const quick = (formula: string) => {
		const started = performance.now();
		const value = calc(formula);
		expect(performance.now() - started, formula).toBeLessThan(1000);
		return value;
	};

	it('answers the patterns that backtrack exponentially in a backtracking matcher', () => {
		expect(quick('REGEXTEST(REPT("a",30),"^(.*a){20}$b")')).toBe(false);
		expect(quick('REGEXTEST(REPT("a",200),".*.*.*.*.*.*.*.*b")')).toBe(false);
		expect(quick('REGEXTEST(REPT("a",30),"^(a*)\\1\\1\\1b")')).toBe(false);
		expect(quick('REGEXTEST(REPT("a",40)&"!","(a+)+$")')).toBe(false);
		expect(quick('REGEXTEST(REPT("a",5000)&"!","(a|aa)+$")')).toBe(false);
		expect(quick('REGEXTEST(REPT("a",300)&"!","(?=(a|aa)+$)")')).toBe(false);
	});

	it('accepts nested quantifiers as Excel does', () => {
		expect(quick('REGEXTEST(REPT("a",40),"^(a+)+$")')).toBe(true);
		expect(quick('REGEXEXTRACT("aaab","(a|aa)+b")')).toBe('aaab');
		expect(quick('REGEXEXTRACT("xaab","(a*)*b",2)')).toBe('aa');
		expect(quick('REGEXREPLACE(REPT("ab",3000),"((a|b)+)+","#")')).toBe('#');
	});
});
