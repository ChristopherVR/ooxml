// Expected values are the worked examples in Microsoft's function reference for each function.
import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers';

const close = (formula: string, expected: number, digits = 5) =>
	expect(calc(formula) as number).toBeCloseTo(expected, digits);

describe('F distribution', () => {
	it('F.DIST left tail, density and right tail', () => {
		close('F.DIST(15.2069,6,4,TRUE)', 0.99, 2);
		close('F.DIST(15.2069,6,4,FALSE)', 0.0012238, 6);
		close('F.DIST.RT(15.2069,6,4)', 0.01, 3);
		close('FDIST(15.2069,6,4)', 0.01, 3);
	});

	it('F.INV and F.INV.RT invert the distribution', () => {
		close('F.INV(0.01,6,4)', 0.10930991, 5);
		close('F.INV.RT(0.01,6,4)', 15.20686, 4);
		close('FINV(0.01,6,4)', 15.20686, 4);
		close('F.DIST(F.INV(0.3,5,7),5,7,TRUE)', 0.3, 8);
	});

	it('rejects invalid degrees of freedom and negative x', () => {
		expect(calc('F.DIST(1,0,4,TRUE)')).toEqual(E.NUM);
		expect(calc('F.DIST(-1,3,4,TRUE)')).toEqual(E.NUM);
		expect(calc('F.INV(1,3,4)')).toEqual(E.NUM);
	});
});

describe('beta, gamma and lognormal inverses and legacy names', () => {
	it('BETA.INV and BETAINV take the optional bounds', () => {
		close('BETA.INV(0.6854706,8,10,1,3)', 2, 5);
		close('BETAINV(0.6854706,8,10,1,3)', 2, 5);
		close('BETADIST(2,8,10,1,3)', 0.6854706, 6);
		close('BETA.INV(0.5,2,3)', calc('BETA.INV(0.5,2,3,0,1)') as number, 12);
	});

	it('GAMMA.INV and GAMMAINV', () => {
		close('GAMMA.INV(0.068094,9,2)', 10.00001191, 5);
		close('GAMMAINV(0.068094,9,2)', 10.00001191, 5);
		close('GAMMADIST(10.00001119,9,2,TRUE)', 0.068094, 5);
		expect(calc('GAMMA.INV(0,9,2)')).toBe(0);
	});

	it('LOGINV and LOGNORMDIST', () => {
		close('LOGINV(0.039084,3.5,1.2)', 4.000025, 5);
		close('LOGNORMDIST(4,3.5,1.2)', 0.0390836, 6);
	});
});

describe('discrete distributions', () => {
	it('BINOM.INV and CRITBINOM', () => {
		expect(calc('BINOM.INV(6,0.5,0.75)')).toBe(4);
		expect(calc('CRITBINOM(6,0.5,0.75)')).toBe(4);
		expect(calc('CRITBINOM(6,0.5,1)')).toBe(6);
		expect(calc('CRITBINOM(6,0.5,2)')).toEqual(E.NUM);
	});

	it('BINOM.DIST.RANGE for a single count and a range', () => {
		close('BINOM.DIST.RANGE(60,0.75,48)', 0.084, 3);
		close('BINOM.DIST.RANGE(60,0.75,45,50)', 0.524, 3);
		expect(calc('BINOM.DIST.RANGE(10,0.5,6,5)')).toEqual(E.NUM);
	});

	it('HYPGEOMDIST and NEGBINOMDIST', () => {
		close('HYPGEOMDIST(1,4,8,20)', 0.363261, 6);
		close('NEGBINOMDIST(10,5,0.25)', 0.055049, 6);
	});
});
