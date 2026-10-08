import { describe, expect, it } from 'vitest';
import { FormulaError } from './ast';
import { formulaShape } from './formula-shape';
import { ParseCache } from './parse-cache';
import { parseFormula } from './parser';

/** What `parseFormula` gives, with errors compared by message. */
function expected(formula: string): unknown {
	try {
		return parseFormula(formula);
	} catch (e) {
		return e instanceof FormulaError ? { error: e.message } : e;
	}
}

function actual(cache: ParseCache, formula: string): unknown {
	const parsed = cache.parse(formula);
	return parsed instanceof FormulaError ? { error: parsed.message } : parsed;
}

/** Pairs of formulas of one shape: the first becomes the template, the second reuses or reparses it. */
const PAIRS: [string, string][] = [
	['B4-C4', 'B5-C5'],
	['MAX(B4-C4,0)*Dashboard!$B$6/365*7', 'MAX(B999-C999,0)*Dashboard!$B$6/365*7'],
	['SUM(A1:A10)*2.5', 'SUM(A2:A11)*2.5'],
	['SUM(A1:A10)*2.5', 'SUM(A2:A11)*3.5'],
	['SUM(1:5)', 'SUM(2:7)'],
	['Sheet2!A1+Sheet2!$B$1', 'Sheet3!A9+Sheet2!$B$1'],
	["'Q1 2024'!B5+1", "'Q1 2025'!B6+1"],
	["'Q1 2024'!B5+1", "'Q1 2024'!B6+1"],
	['[1]Sheet1!A1', '[2]Sheet1!A2'],
	['LOG10(A1)', 'LOG10(A2)'],
	['LOG10(A1)', 'LOG2(A2)'],
	['"x1"&A1', '"x2"&A2'],
	['"x1"&A1', '"x1"&A2'],
	['A1#', 'A2#'],
	['A01+1', 'A02+1'],
	['A1+1', 'A01+1'],
	['A1048576', 'A1048577'],
	['A0+1', 'A1+1'],
	['Rate2*A1', 'Rate3*A2'],
	['1E5+A1', '1E6+A2'],
	['SUM(Table1[Col1])+A1', 'SUM(Table1[Col1])+A2'],
	['{1,2;3,4}', '{1,2;3,5}'],
	['A1:B2 B2:C3', 'A5:B6 B6:C7'],
	['INDEX(A1:A9,2)', 'INDEX(A2:A10,2)'],
	['XFD1+AAA2', 'XFD3+AAA4'],
	['#REF!+Sheet1!#REF!+A1', '#REF!+Sheet1!#REF!+A2'],
	['SUM(A1', 'SUM(A2'],
];

describe('ParseCache', () => {
	it('gives the tree parseFormula gives for every copy of a shape', () => {
		for (const [first, second] of PAIRS) {
			const cache = new ParseCache();
			expect(actual(cache, first), first).toEqual(expected(first));
			expect(actual(cache, second), second).toEqual(expected(second));
		}
	});

	it('reuses a template for copies filled down a column', () => {
		const cache = new ParseCache();
		const trees = Array.from({ length: 50 }, (_, r) =>
			cache.parse(`IF(A${r + 1}>0,B${r + 1}*Rates!$C$2,0)`),
		);
		trees.forEach((tree, r) => {
			expect(tree).toEqual(parseFormula(`IF(A${r + 1}>0,B${r + 1}*Rates!$C$2,0)`));
		});
		// Untouched parts of the tree (the absolute reference) are shared with the template.
		const rate = (tree: unknown) => (tree as { args: { right?: unknown }[] }).args[1]?.right;
		expect(rate(trees[7])).toMatchObject({ type: 'ref' });
		expect(rate(trees[7])).toBe(rate(trees[0]));
	});

	it('returns the same tree for the same text', () => {
		const cache = new ParseCache();
		expect(cache.parse('A1+B1')).toBe(cache.parse('A1+B1'));
	});

	it('blanks digit runs in the shape', () => {
		const a = formulaShape('B12-C12+1.5');
		const b = formulaShape('B7-C7+3.25');
		expect(a.hash).toBe(b.hash);
		expect(a.starts).toEqual([1, 5, 8, 10]);
		expect(formulaShape('B1-C1').hash).not.toBe(formulaShape('B1+C1').hash);
	});
});
