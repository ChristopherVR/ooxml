import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers';

const near = (formula: string, expected: number, digits = 8) =>
	expect(calc(formula) as number).toBeCloseTo(expected, digits);

describe('Bessel functions', () => {
	it('matches the documented examples', () => {
		near('BESSELI(1.5,1)', 0.981666428, 8);
		// Microsoft documents 0.329925829 (Excel's own series error, ~1e-7); the exact value is 0.3299257277.
		near('BESSELJ(1.9,2)', 0.329925829, 6);
		near('BESSELK(1.5,1)', 0.277387804, 8);
		near('BESSELY(2.5,1)', 0.145918138, 8);
	});
	it('matches reference values', () => {
		near('BESSELJ(1,0)', 0.7651976866, 9);
		near('BESSELJ(10,3)', 0.05837937933, 9);
		near('BESSELY(1,0)', 0.08825696422, 9);
		near('BESSELY(10,0)', 0.05567116728, 9);
		near('BESSELI(1,0)', 1.2660658778, 9);
		near('BESSELK(1,0)', 0.4210244382, 9);
		near('BESSELK(10,0)', 1.778006232e-5, 12);
	});
	it('handles signs and zero', () => {
		near('BESSELJ(-2,1)', -0.5767248078, 9);
		expect(calc('BESSELJ(0,0)')).toBe(1);
		expect(calc('BESSELI(0,2)')).toBe(0);
	});
	it('reports domain errors', () => {
		expect(calc('BESSELJ(1,-1)')).toEqual(E.NUM);
		expect(calc('BESSELY(0,1)')).toEqual(E.NUM);
		expect(calc('BESSELK(-1,1)')).toEqual(E.NUM);
		expect(calc('BESSELJ("a",1)')).toEqual(E.VALUE);
	});
});
