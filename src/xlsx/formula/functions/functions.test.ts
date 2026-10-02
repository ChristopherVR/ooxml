import { describe, expect, it } from 'vitest';
import { calc, calcArray, E } from '../test-helpers.js';
import { FUNCTION_CATALOG, getFunction } from './registry.js';

describe('FUNCTION_CATALOG', () => {
	it('lists well over 150 implemented functions with unique names', () => {
		const names = FUNCTION_CATALOG.map((f) => f.name);
		expect(names.length).toBeGreaterThanOrEqual(150);
		expect(new Set(names).size).toBe(names.length);
	});

	it('documents every function', () => {
		for (const info of FUNCTION_CATALOG) {
			expect(info.syntax.startsWith(`${info.name}(`)).toBe(true);
			expect(info.description.length).toBeGreaterThan(5);
			expect(info.description).not.toContain(String.fromCharCode(0x2014));
			expect(getFunction(info.name)).toBeDefined();
		}
	});

	it('organises functions by category', () => {
		const categories = new Set(FUNCTION_CATALOG.map((f) => f.category));
		for (const c of [
			'Math & Trig',
			'Statistical',
			'Logical',
			'Text',
			'Date & Time',
			'Lookup & Reference',
			'Information',
			'Financial',
			'Engineering',
		]) {
			expect(categories).toContain(c);
		}
	});
});

describe('dynamic array functions', () => {
	const data = { A1: 3, A2: 1, A3: 2, B1: 'c', B2: 'a', B3: 'b' };

	it('SORT and SORTBY spill sorted rows', () => {
		expect(calcArray('SORT(A1:B3)', data, 3, 2)).toEqual([
			[1, 'a'],
			[2, 'b'],
			[3, 'c'],
		]);
		expect(calcArray('SORT(A1:B3,2,-1)', data, 3, 2)).toEqual([
			[3, 'c'],
			[2, 'b'],
			[1, 'a'],
		]);
		expect(calcArray('SORTBY(B1:B3,A1:A3)', data, 3, 1)).toEqual([['a'], ['b'], ['c']]);
	});

	it('FILTER keeps matching rows and reports #CALC! when nothing matches', () => {
		expect(calcArray('FILTER(A1:B3,A1:A3>1)', data, 2, 2)).toEqual([
			[3, 'c'],
			[2, 'b'],
		]);
		expect(calc('FILTER(A1:A3,A1:A3>5)', data)).toEqual(E.CALC);
		expect(calc('FILTER(A1:A3,{1,0})', data)).toEqual(E.VALUE);
	});

	it('UNIQUE compares text case-insensitively', () => {
		expect(calcArray('UNIQUE({"a";"A";"b";"a"})', {}, 3, 1)).toEqual([['a'], ['b'], [null]]);
		expect(calcArray('UNIQUE({1,1,2},TRUE)', {}, 1, 3)).toEqual([[1, 2, null]]);
	});

	it('SEQUENCE, RANDARRAY and MAKEARRAY produce the requested shape', () => {
		expect(calcArray('SEQUENCE(2,2,0,-1)', {}, 2, 2)).toEqual([
			[0, -1],
			[-2, -3],
		]);
		expect(calc('SEQUENCE(0)')).toEqual(E.CALC);
		const random = calcArray('RANDARRAY(2,2,1,6,TRUE)', {}, 2, 2).flat();
		expect(
			random.every((v) => typeof v === 'number' && Number.isInteger(v) && v >= 1 && v <= 6),
		).toBe(true);
		expect(calcArray('MAKEARRAY(2,2,LAMBDA(r,c,r&c))', {}, 2, 2)).toEqual([
			['11', '12'],
			['21', '22'],
		]);
	});

	it('XLOOKUP returns whole rows from a 2D return array', () => {
		expect(calcArray('XLOOKUP("b",B1:B3,A1:B3)', data, 1, 2)).toEqual([[2, 'b']]);
		expect(calc('XLOOKUP(1,A1:A3,B1:B2)', data)).toEqual(E.VALUE);
	});

	it('TRANSPOSE, TOCOL and TEXTSPLIT pad and reshape', () => {
		expect(calcArray('TRANSPOSE(A1:A3)', data, 1, 3)).toEqual([[3, 1, 2]]);
		expect(calcArray('TEXTSPLIT("a,b;c",",",";")', {}, 2, 2)).toEqual([
			['a', 'b'],
			['c', E.NA],
		]);
		expect(calcArray('TOCOL({1,2;3,4},0,TRUE)', {}, 4, 1)).toEqual([[1], [3], [2], [4]]);
	});

	it('FREQUENCY and MODE.MULT return vertical arrays', () => {
		expect(calcArray('FREQUENCY({1,2,3,4,5},{2,4})', {}, 3, 1)).toEqual([[2], [2], [1]]);
		expect(calcArray('MODE.MULT({1,1,2,2,3})', {}, 2, 1)).toEqual([[1], [2]]);
	});

	it('ROW and COLUMN spill for multi-cell references', () => {
		expect(calcArray('ROW(B2:B4)', {}, 3, 1)).toEqual([[2], [3], [4]]);
		expect(calcArray('COLUMN(B2:D2)', {}, 1, 3)).toEqual([[2, 3, 4]]);
	});

	it('MMULT and MINVERSE spill matrices', () => {
		expect(calcArray('MMULT({1,2;3,4},{1,0;0,1})', {}, 2, 2)).toEqual([
			[1, 2],
			[3, 4],
		]);
		expect(calc('MINVERSE({1,2;2,4})')).toEqual(E.NUM);
		expect(calc('MMULT({1,2},{1,2})')).toEqual(E.VALUE);
	});
});

describe('argument handling edge cases', () => {
	it('distinguishes omitted and empty optional arguments', () => {
		expect(calc('VLOOKUP(2,{1,"a";3,"b"},2)')).toBe('a');
		expect(calc('VLOOKUP(2,{1,"a";3,"b"},2,)')).toEqual(E.NA);
		expect(calc('ROUND(2.567,)')).toBe(3);
	});

	it('criteria match numbers, numeric text, blanks and wildcards', () => {
		const cells = { A1: 5, A2: '5', A3: 'x5', A5: 'a~b', A6: 'a*b' };
		expect(calc('COUNTIF(A1:A6,5)', cells)).toBe(2);
		expect(calc('COUNTIF(A1:A6,"=")', cells)).toBe(1);
		expect(calc('COUNTIF(A1:A6,"*5")', cells)).toBe(2);
		expect(calc('COUNTIF(A1:A6,"a~*b")', cells)).toBe(1);
		expect(calc('COUNTIF(A1:A6,">4")', cells)).toBe(1);
		expect(calc('SUMIF(A1:A6,"<>x5")', cells)).toBe(5);
	});

	it('SUMIF resizes the sum range from its top-left cell', () => {
		expect(calc('SUMIF(A1:A3,">0",B1)', { A1: 1, A2: 0, A3: 2, B1: 10, B2: 20, B3: 30 })).toBe(40);
		expect(calc('SUMIFS(B1:B3,A1:A2,">0")', { A1: 1, B1: 1 })).toEqual(E.VALUE);
	});

	it('SUBTOTAL ignores nested subtotals', () => {
		expect(calc('SUBTOTAL(9,A1:A3)', { A1: 1, A2: 2, A3: '=SUBTOTAL(9,A1:A2)' })).toBe(3);
	});

	it('aggregates ignore text and logicals in ranges but coerce direct arguments', () => {
		expect(calc('SUM(A1:A3)', { A1: 1, A2: '2', A3: true })).toBe(1);
		expect(calc('SUM(1,"2",TRUE)')).toBe(4);
		expect(calc('SUM("x")')).toEqual(E.VALUE);
		expect(calc('AVERAGE(A1:A2)', { A1: 'a' })).toEqual(E.DIV0);
		expect(calc('MAX(A1:A2)', { A1: 'a' })).toBe(0);
	});

	it('IF evaluates only the chosen branch', () => {
		expect(calc('IF(TRUE,1,NOSUCH())')).toBe(1);
		expect(calc('IFERROR(1/0,NOSUCH())')).toEqual(E.NAME);
		expect(calcArray('IF({1,0},"y","n")', {}, 1, 2)).toEqual([['y', 'n']]);
	});

	it('text functions validate their arguments', () => {
		expect(calc('LEFT("abc",-1)')).toEqual(E.VALUE);
		expect(calc('REPT("x",40000)')).toEqual(E.VALUE);
		expect(calc('CHAR(0)')).toEqual(E.VALUE);
		expect(calc('FIND("z","abc")')).toEqual(E.VALUE);
		expect(calc('CODE("")')).toEqual(E.VALUE);
	});

	it('date functions reject out-of-range serials', () => {
		expect(calc('DAY(-1)')).toEqual(E.NUM);
		expect(calc('YEAR(3000000)')).toEqual(E.NUM);
		expect(calc('DATEDIF(2,1,"D")')).toEqual(E.NUM);
		expect(calc('WEEKDAY(1,9)')).toEqual(E.NUM);
	});

	it('financial solvers report #NUM! without a sign change', () => {
		expect(calc('IRR({1,2,3})')).toEqual(E.NUM);
		expect(calc('PMT(0.1,0,100)')).toEqual(E.NUM);
	});

	it('engineering conversions validate digits and widths', () => {
		expect(calc('BIN2DEC("102")')).toEqual(E.NUM);
		expect(calc('DEC2BIN(512)')).toEqual(E.NUM);
		expect(calc('DEC2BIN(3,1)')).toEqual(E.NUM);
		expect(calc('HEX2OCT("1F",4)')).toBe('0037');
	});
});
