import { describe, expect, it } from 'vitest';
import {
	referencedRanges,
	renameSheetInFormula,
	shiftFormula,
	translateFormula,
} from './transform.js';

describe('translateFormula', () => {
	it.each([
		['A1+$B$2+B$3+$C4', 1, 1, 'B2+$B$2+C$3+$C5'],
		['SUM(A1:B2)', 2, 0, 'SUM(A3:B4)'],
		['SUM( a1 , 10 )', 0, 1, 'SUM( B1 , 10 )'],
		['A:A', 5, 2, 'C:C'],
		['$A:B', 0, 1, '$A:C'],
		['3:4', 2, 9, '5:6'],
		["Sheet2!A1*'My Sheet'!B2", 1, 0, "Sheet2!A2*'My Sheet'!B3"],
		['A1#', 1, 1, 'B2#'],
		['"A1"&A1', 1, 0, '"A1"&A2'],
		['Rate*Table1[Qty]', 3, 3, 'Rate*Table1[Qty]'],
		['=A1', 1, 0, '=A2'],
	])('%s moved by (%i,%i) is %s', (formula, dRow, dCol, expected) => {
		expect(translateFormula(formula, dRow, dCol)).toBe(expected);
	});

	it('turns references moved off the grid into #REF!', () => {
		expect(translateFormula('A1+B2', -1, 0)).toBe('#REF!+B1');
		expect(translateFormula('Sheet2!A1:B2', 0, -1)).toBe('Sheet2!#REF!');
		expect(translateFormula('XFD1', 0, 1)).toBe('#REF!');
	});

	it('keeps text it cannot tokenize unchanged', () => {
		expect(translateFormula('"unterminated', 1, 1)).toBe('"unterminated');
	});
});

describe('shiftFormula', () => {
	const rows = (at: number, count: number, sheet = 'Sheet1') => ({
		sheet,
		axis: 'row' as const,
		at,
		count,
	});
	const cols = (at: number, count: number, sheet = 'Sheet1') => ({
		sheet,
		axis: 'col' as const,
		at,
		count,
	});

	it.each([
		['A1+A5', rows(2, 2), 'A1+A7'],
		['$A$5', rows(0, 1), '$A$6'],
		['SUM(A1:A10)', rows(4, 3), 'SUM(A1:A13)'],
		['SUM(A1:A5)', rows(5, 1), 'SUM(A1:A5)'],
		['SUM(A3:A5)', rows(2, 1), 'SUM(A4:A6)'],
		['A:A', rows(0, 5), 'A:A'],
		['2:3', rows(0, 1), '3:4'],
		['C1+D1', cols(2, 1), 'D1+E1'],
		['B:D', cols(0, 2), 'D:F'],
		['1:1', cols(0, 3), '1:1'],
	])('insert: %s', (formula, spec, expected) => {
		expect(shiftFormula(formula, 'Sheet1', spec)).toBe(expected);
	});

	it.each([
		['A1+A5', rows(1, 2), 'A1+A3'],
		['A3', rows(2, 1), '#REF!'],
		['SUM(A1:A10)', rows(2, 3), 'SUM(A1:A7)'],
		['SUM(A3:A5)', rows(1, 4), 'SUM(#REF!)'],
		['SUM(A3:A8)', rows(1, 4), 'SUM(A2:A4)'],
		['SUM(A3:A8)', rows(5, 10), 'SUM(A3:A5)'],
		['Sheet1!B2:C3', cols(1, 1), 'Sheet1!B2:B3'],
		['B:B', cols(1, 1), '#REF!'],
		['5:9', rows(0, 2), '3:7'],
	])('delete: %s', (formula, spec, expected) => {
		expect(shiftFormula(formula, 'Sheet1', { ...spec, count: -spec.count })).toBe(expected);
	});

	it('only shifts references to the edited sheet', () => {
		expect(shiftFormula('A5+Sheet2!A5', 'Sheet1', rows(0, 1, 'Sheet2'))).toBe('A5+Sheet2!A6');
		expect(shiftFormula('A5', 'Sheet2', rows(0, 1, 'sheet2'))).toBe('A6');
		expect(shiftFormula("'Data Set'!A5", 'Sheet1', rows(0, 1, 'Data Set'))).toBe("'Data Set'!A6");
	});

	it('preserves surrounding text and whitespace', () => {
		expect(shiftFormula('IF( A5 > 0 , "A5" , B5 )', 'Sheet1', rows(0, 1))).toBe(
			'IF( A6 > 0 , "A5" , B6 )',
		);
	});
});

describe('renameSheetInFormula', () => {
	it.each([
		['Sheet1!A1+Sheet2!B2', 'Sheet1', 'Data', 'Data!A1+Sheet2!B2'],
		['sheet1!A1', 'Sheet1', 'My Data', "'My Data'!A1"],
		["'Old Name'!A1:B2", 'Old Name', 'New', 'New!A1:B2'],
		['Sheet1!Rate*2', 'Sheet1', "Bob's", "'Bob''s'!Rate*2"],
		['Sheet1:Sheet3!A1', 'Sheet3', 'End', 'Sheet1:End!A1'],
		['SUM(Sheet1!A:A)', 'Other', 'X', 'SUM(Sheet1!A:A)'],
		['Sheet1!A1', 'Sheet1', 'A1', "'A1'!A1"],
		['INDIRECT("Sheet1!A1")', 'Sheet1', 'X', 'INDIRECT("Sheet1!A1")'],
	])('%s: %s -> %s', (formula, from, to, expected) => {
		expect(renameSheetInFormula(formula, from, to)).toBe(expected);
	});
});

describe('referencedRanges', () => {
	it('lists the ranges a formula refers to', () => {
		expect(referencedRanges("SUM(A1:B2)+'My Sheet'!C3*D:D", 'Sheet1')).toEqual([
			{ sheet: 'Sheet1', range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } } },
			{ sheet: 'My Sheet', range: { start: { row: 2, col: 2 }, end: { row: 2, col: 2 } } },
			{ sheet: 'Sheet1', range: { start: { row: 0, col: 3 }, end: { row: 1_048_575, col: 3 } } },
		]);
	});

	it('normalizes reversed ranges and skips names, #REF! and external refs', () => {
		expect(referencedRanges('B2:A1+Rate+#REF!+[1]S!A1', 'S')).toEqual([
			{ sheet: 'S', range: { start: { row: 0, col: 0 }, end: { row: 1, col: 1 } } },
		]);
		expect(referencedRanges('"broken', 'S')).toEqual([]);
	});
});
