import { describe, expect, it } from 'vitest';
import { analyzeVisioFormula, evaluateVisioFormula, type VisioFormulaValue } from './formula';

// Expected values were recorded from Visio 16 (scripts/record-visio-formula-functions.ps1).
const resolve = (reference: { cell: string }): VisioFormulaValue => {
	if (reference.cell === 'Width') return { value: 3, unit: 'length' };
	if (reference.cell === 'Name') throw new Error('not numeric');
	throw new Error(`Unknown cell ${reference.cell}`);
};
const run = (source: string) => evaluateVisioFormula(source, resolve);
const number = (source: string) => run(source).value;

describe('ShapeSheet string functions on literal text', () => {
	it('measures and searches text', () => {
		expect(number('LEN("hello")')).toBe(5);
		expect(number('FIND("l","hello")')).toBe(3);
		expect(number('FIND("l","hello",4)')).toBe(4);
		expect(number('FIND("","hello")')).toBe(1);
		expect(() => run('FIND("z","hello")')).toThrow(/did not find/);
		expect(() => run('FIND("L","hello")')).toThrow(/did not find/);
		expect(() => run('FIND("l","hello",0)')).toThrow(/outside/);
		expect(() => run('FIND("l","hello",6)')).toThrow(/outside/);
	});

	it('compares text with STRSAME and finds list items with LOOKUP', () => {
		expect(number('STRSAME("a","a")')).toBe(1);
		expect(number('STRSAME("abc","ABC")')).toBe(0);
		expect(number('STRSAME("abc","ABC",1)')).toBe(1);
		expect(number('STRSAME("a","b",2)')).toBe(0);
		expect(number('LOOKUP("b","a;b;c")')).toBe(1);
		expect(number('LOOKUP("B","a;b;c")')).toBe(1);
		expect(number('LOOKUP("z","a;b;c")')).toBe(-1);
		expect(number('LOOKUP("b","a;;b")')).toBe(2);
		expect(number('LOOKUP("","a;;b")')).toBe(1);
	});

	it('takes parts of text', () => {
		expect(number('LEN(MID("hello",2,3))')).toBe(3);
		expect(number('LEN(MID("hello",4,100))')).toBe(2);
		expect(number('LEN(MID("hello",9,1))')).toBe(0);
		expect(number('LEN(MID("hello",2,-1))')).toBe(4);
		expect(() => run('LEN(MID("hello",0,3))')).toThrow(/at least 1/);
		expect(number('LEN(LEFT("hello",2))')).toBe(2);
		expect(number('LEN(LEFT("hello"))')).toBe(1);
		expect(number('LEN(LEFT("hello",0))')).toBe(0);
		expect(number('LEN(LEFT("hello",-1))')).toBe(5);
		expect(number('LEN(RIGHT("hello",9))')).toBe(5);
		expect(number('LEN(RIGHT("hello"))')).toBe(1);
		expect(number('FIND("l",MID("hello",4,2))')).toBe(1);
	});

	it('changes text', () => {
		expect(number('LEN(UPPER("ab")&"cd")')).toBe(4);
		expect(number('LEN(LOWER("ABC"))')).toBe(3);
		expect(number('LEN(TRIM("  a  b "))')).toBe(3);
		expect(number('LEN(SUBSTITUTE("aXbX","X","yy"))')).toBe(6);
		expect(number('LEN(SUBSTITUTE("aaa","a","bb",2))')).toBe(4);
		expect(number('LEN(SUBSTITUTE("aaa","","b"))')).toBe(3);
		expect(number('LEN(REPLACE("hello",2,2,"ZZZZ"))')).toBe(7);
		expect(number('LEN(REPLACE("hello",4,10,"Z"))')).toBe(4);
		expect(() => run('LEN(REPLACE("hello",0,2,"Z"))')).toThrow(/at least 1/);
		expect(number('LEN(CHAR(65))')).toBe(1);
		expect(number('FIND("A",CHAR(65))')).toBe(1);
		expect(() => run('LEN(CHAR(0))')).toThrow(/1 to 255/);
		expect(() => run('LEN(CHAR(300))')).toThrow(/1 to 255/);
	});

	it('picks list items with INDEX and joins numbers with &', () => {
		expect(number('LEN(INDEX(0,"abc;d"))')).toBe(3);
		expect(number('LEN(INDEX(2.7,"a;bb;ccc"))')).toBe(3);
		expect(number('LEN(INDEX(5,"a;b;c"))')).toBe(0);
		expect(number('LEN(INDEX(-1,"a;b;c"))')).toBe(0);
		expect(number('LEN("a"&5)')).toBe(2);
		expect(number('LEN("a"&1.5)')).toBe(4);
		expect(number('LEN(5)')).toBe(1);
		expect(number('LEN(1.5)')).toBe(3);
	});

	it('reports ISERROR and IFERROR as the sheet does', () => {
		expect(number('ISERROR(FIND("z","hello"))')).toBe(1);
		expect(number('ISERROR(FIND("l","hello"))')).toBe(0);
		expect(number('IFERROR(FIND("z","hello"),-5)')).toBe(-5);
		expect(number('IFERROR(FIND("l","hello"),-5)')).toBe(3);
		expect(number('IFERROR(1/0,7)')).toBe(7);
		// A function this evaluator cannot run is not an error of the sheet.
		expect(() => run('IFERROR(STRSAMEEX("a","a"),1)')).toThrow(/Unsupported/);
	});

	it('refuses text it cannot know and text results', () => {
		expect(() => run('LEN(Name)')).toThrow(/Text read from a cell/);
		expect(() => run('MID("abc",1,1)')).toThrow(/does not return a number/);
		expect(() => run('"abc"')).toThrow(/String results/);
		expect(() => run('LEN("a"&Width)')).toThrow(/unit/);
		expect(() => run('LEN()')).toThrow(/argument count/);
		expect(() => run('FIND("a")')).toThrow(/argument count/);
	});

	it('classifies the functions as supported', () => {
		expect(
			analyzeVisioFormula('LEN(MID("hello",2,3))+FIND("l","hello")').unsupportedFunctions,
		).toEqual([]);
		expect(analyzeVisioFormula('STRSAMEEX("a","a",1,1)').unsupportedFunctions).toEqual([
			'STRSAMEEX',
		]);
	});
});

describe('ShapeSheet angle, sum and point functions', () => {
	it('labels and converts angles as Visio does', () => {
		expect(run('DEG(90)')).toEqual({ value: 90, unit: 'angle' });
		expect(run('DEG(90 deg)').value).toBeCloseTo(Math.PI / 2, 12);
		expect(run('RAD(90)')).toEqual({ value: Math.PI / 2, unit: 'angle' });
		expect(run('RAD(90 deg)').value).toBeCloseTo(Math.PI / 2, 12);
		expect(run('RAD(180)').value).toBeCloseTo(Math.PI, 12);
		expect(() => run('RAD(Width)')).toThrow(/angle or a number/);
	});

	it('wraps an angle into a turn with ANG360', () => {
		expect(run('ANG360(-90 deg)').value).toBeCloseTo((3 * Math.PI) / 2, 12);
		expect(run('ANG360(450 deg)').value).toBeCloseTo(Math.PI / 2, 12);
		expect(run('ANG360(0)')).toEqual({ value: 0, unit: 'angle' });
		expect(run('ANG360(360 deg)')).toEqual({ value: 0, unit: 'angle' });
		expect(run('ANG360(-0.5 deg)').value).toBeCloseTo(2 * Math.PI - (0.5 * Math.PI) / 180, 12);
	});

	it('adds with SUM and reads PNT coordinates', () => {
		expect(run('SUM(1,2,3)')).toEqual({ value: 6, unit: 'scalar' });
		expect(run('SUM(1 in,2 in)')).toEqual({ value: 3, unit: 'length' });
		expect(() => run('SUM(1 in,2)')).toThrow(/Incompatible/);
		expect(run('PNTX(PNT(3,4))')).toEqual({ value: 3, unit: 'length' });
		expect(run('PNTY(PNT(3,4))')).toEqual({ value: 4, unit: 'length' });
		expect(run('PNTX(PNT(Width*2,4 mm))').value).toBe(6);
		expect(() => run('PNT(3,4)')).toThrow(/does not return a number/);
		expect(() => run('PNTX(3)')).toThrow(/PNT/);
	});

	it('still refuses what Visio documents but this evaluator does not run', () => {
		for (const name of ['STRSAMEEX', 'ANGLETOLOC', 'LOCTOPAR', 'FORMAT', 'DATE', 'INDIRECT'])
			expect(() => run(`${name}(1)`)).toThrow();
	});
});
