import { describe, expect, it } from 'vitest';
import { createCalcEngine } from './engine';
import { compileWildcard } from './functions/wildcard';
import { book, calc, E, engine, get } from './test-helpers';
import { CalculationStopped, MAX_STORED_CELLS, RecalcLimits } from './recalc-limits';
import { chargeWork, Matrix, MAX_FORMULA_WORK, resetCalcBudget, withFormulaBudget } from './values';

describe('defined name work limits', () => {
	it('evaluates a 31-level doubling chain quickly', () => {
		const wb = book({ A1: '=Dbl_31' });
		wb.definedNames.push({ name: 'Dbl_0', formula: '1' });
		for (let i = 1; i <= 31; i++)
			wb.definedNames.push({ name: `Dbl_${i}`, formula: `Dbl_${i - 1}+Dbl_${i - 1}` });
		const started = Date.now();
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toBe(2 ** 31);
		expect(Date.now() - started).toBeLessThan(5000);
	});

	it('returns an error for a self-referencing name', () => {
		const wb = book({ A1: '=Loop' });
		wb.definedNames.push({ name: 'Loop', formula: 'Loop+Loop' });
		const started = Date.now();
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.NAME);
		expect(Date.now() - started).toBeLessThan(5000);
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
		expect(Date.now() - started).toBeLessThan(5000);
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

describe('formula array budget', () => {
	it('fails once a formula has built too many cells', () => {
		resetCalcBudget();
		try {
			for (let i = 0; i < 4; i++) Matrix.build(4000, 3000, () => 0);
			expect(() => Matrix.build(4000, 3000, () => 0)).toThrow();
		} finally {
			resetCalcBudget();
		}
	});
});

describe('formula work and text budget', () => {
	it('stops a cubic matrix product, elimination and text growth with #NUM!', () => {
		expect(calc('SUM(MMULT(SEQUENCE(1500,1500),SEQUENCE(1500,1500)))')).toEqual(E.NUM);
		expect(calc('MDETERM(SEQUENCE(1000,1000))')).toEqual(E.NUM);
		expect(calc('SUM(LEN(MAKEARRAY(2000,2000,LAMBDA(r,c,REPT("a",32767)&r&c))))')).toEqual(E.NUM);
	});

	it('still computes work of an ordinary size', () => {
		expect(calc('SUM(MMULT(SEQUENCE(300,300),SEQUENCE(300,300)))')).toBe(
			calc('SUMPRODUCT(MMULT(SEQUENCE(300,300),SEQUENCE(300,300)))'),
		);
		expect(calc('REDUCE(0,SEQUENCE(300,300),LAMBDA(a,b,a+b))')).toBe(4050045000);
		expect(calc('LEN(CONCAT(MAKEARRAY(100,100,LAMBDA(r,c,REPT("a",3)))))')).toBe(30000);
	});

	it('gives every formula its own budget, so a refused formula never refuses the next', () => {
		const wb = book({
			A1: '=SUM(MMULT(SEQUENCE(1500,1500),SEQUENCE(1500,1500)))',
			A2: '=SUM(MMULT(SEQUENCE(300,300),SEQUENCE(300,300)))',
			A4: '=REDUCE(0,SEQUENCE(300,300),LAMBDA(a,b,a+b))',
		});
		engine(wb).recalculateAll();
		expect(get(wb, 'A1')).toEqual(E.NUM);
		expect(get(wb, 'A2')).toBe(calc('SUMPRODUCT(MMULT(SEQUENCE(300,300),SEQUENCE(300,300)))'));
		expect(get(wb, 'A4')).toBe(4050045000);
	});

	it('keeps what a formula has left while a formula it reads is computed', () => {
		withFormulaBudget(() => {
			chargeWork(MAX_FORMULA_WORK - 10);
			withFormulaBudget(() => chargeWork(MAX_FORMULA_WORK - 10));
			expect(() => chargeWork(20)).toThrow();
		});
		expect(() => chargeWork(MAX_FORMULA_WORK - 10)).not.toThrow();
		resetCalcBudget();
	});

	it('gives every recalculation a fresh budget', () => {
		const wb = book({ A1: '=SUM(MMULT(SEQUENCE(500,500),SEQUENCE(500,500)))' });
		const calcEngine = engine(wb);
		for (let i = 0; i < 6; i++) {
			calcEngine.recalculateAll();
			expect(typeof get(wb, 'A1')).toBe('number');
		}
	});
});

describe('scans an array of lookups or criteria repeats', () => {
	const spill = (formula: string, n: number) => {
		const wb = book({ A1: `=SEQUENCE(${n})`, C1: `=${formula}` });
		engine(wb).recalculateAll();
		return get(wb, 'C1');
	};

	it('stops a criteria or nearest-match scan repeated over a long array', () => {
		expect(spill('SUM(COUNTIF(A1#,">"&SEQUENCE(20000)))', 20000)).toEqual(E.NUM);
		expect(spill('SUM(XLOOKUP(SEQUENCE(20000)-0.5,A1#,A1#,,1))', 20000)).toEqual(E.NUM);
	});

	it('still answers short repeated scans and indexed exact lookups', () => {
		expect(spill('SUM(COUNTIF(A1#,">"&SEQUENCE(200)))', 200)).toBe(19900);
		expect(spill('SUM(XMATCH(SEQUENCE(20000),A1#))', 20000)).toBe(200010000);
		expect(spill('SUM(MATCH(SEQUENCE(20000)+0.5,A1#,1))', 20000)).toBe(200010000);
	});
});

describe('recalculation limits', () => {
	it('stops a recalculation whose results would hold too many cells or too much text', () => {
		const limits = new RecalcLimits(undefined);
		limits.begin();
		const cells = Matrix.padded(new Matrix([[0]]), 10_000, 4_000, 0);
		limits.store(cells, 40_000_000);
		expect(() => limits.store(cells, 40_000_000)).toThrow(CalculationStopped);
		limits.begin();
		limits.store(cells, MAX_STORED_CELLS);
		limits.begin();
		const text = Matrix.padded(new Matrix([['x']]), 1_000_000, 300, 'xx');
		expect(() => limits.store(text, 1)).toThrow(CalculationStopped);
	});

	it('stops, rather than failing the formula, at one that needs more work than a cap', () => {
		const wb = book({ A1: '=1+1', A2: '=SUM(MMULT(SEQUENCE(300,300),SEQUENCE(300,300)))' });
		const capped = createCalcEngine(wb, { workCap: 1_000_000 });
		capped.recalculateAll();
		expect(capped.stopped).toBe('work');
		expect(get(wb, 'A2')).toBeNull();
		const full = createCalcEngine(wb);
		full.recalculateAll();
		expect(full.stopped).toBeUndefined();
		expect(typeof get(wb, 'A2')).toBe('number');
	});

	it('stops at the time limit between formulas and computes everything next time', () => {
		const cells: Record<string, string> = {};
		for (let row = 1; row <= 200; row++) cells[`A${row}`] = '=ROW()';
		const wb = book(cells);
		const limited = createCalcEngine(wb, { timeLimitMs: 0 });
		limited.recalculateAll();
		expect(limited.stopped).toBe('time');
		expect(get(wb, 'A200')).toBeNull();
		limited.recalculateAll();
		expect(limited.stopped).toBe('time');
		const unlimited = createCalcEngine(wb);
		unlimited.recalculateAll();
		expect(unlimited.stopped).toBeUndefined();
		expect(get(wb, 'A200')).toBe(200);
	});
});
