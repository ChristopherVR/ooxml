import { describe, expect, it } from 'vitest';
import { computeLinearRegression, computeRSquared, fitPolynomial } from './regression.js';

describe('computeLinearRegression', () => {
	it('returns slope=1 intercept=0 for y=x data', () => {
		const xs = [0, 1, 2, 3];
		const ys = [0, 1, 2, 3];
		const { slope, intercept, rSquared } = computeLinearRegression(xs, ys);
		expect(slope).toBeCloseTo(1, 6);
		expect(intercept).toBeCloseTo(0, 6);
		expect(rSquared).toBeCloseTo(1, 6);
	});

	it('returns slope=2 intercept=1 for y=2x+1 data', () => {
		const xs = [0, 1, 2, 3, 4];
		const ys = xs.map((x) => 2 * x + 1);
		const { slope, intercept, rSquared } = computeLinearRegression(xs, ys);
		expect(slope).toBeCloseTo(2, 5);
		expect(intercept).toBeCloseTo(1, 5);
		expect(rSquared).toBeCloseTo(1, 5);
	});

	it('returns rSquared < 1 for noisy data', () => {
		const xs = [0, 1, 2, 3, 4];
		const ys = [0, 2, 1, 4, 3];
		const { rSquared } = computeLinearRegression(xs, ys);
		expect(rSquared).toBeGreaterThan(0);
		expect(rSquared).toBeLessThan(1);
	});

	it('returns zeros for fewer than 2 points', () => {
		const result = computeLinearRegression([1], [1]);
		expect(result.slope).toBe(0);
		expect(result.intercept).toBe(0);
		expect(result.rSquared).toBe(0);
	});

	it('handles zero denominator (all x equal)', () => {
		const result = computeLinearRegression([2, 2, 2], [1, 2, 3]);
		expect(result.slope).toBe(0);
		expect(result.intercept).toBeCloseTo(2, 6); // mean of y
		expect(result.rSquared).toBe(0);
	});

	it('returns zero for empty arrays', () => {
		const result = computeLinearRegression([], []);
		expect(result.slope).toBe(0);
		expect(result.intercept).toBe(0);
		expect(result.rSquared).toBe(0);
	});
});

// ─────────────────────────────────────────────────────────────────────────────
// fitPolynomial
// ─────────────────────────────────────────────────────────────────────────────

describe('fitPolynomial', () => {
	it('recovers linear coefficients for degree-1 fit', () => {
		const xs = [0, 1, 2, 3];
		const ys = xs.map((x) => 3 * x + 5);
		const coeffs = fitPolynomial(xs, ys, 1);
		// coeffs[0] = intercept, coeffs[1] = slope
		expect(coeffs[0]).toBeCloseTo(5, 4);
		expect(coeffs[1]).toBeCloseTo(3, 4);
	});

	it('recovers quadratic coefficients for degree-2 fit', () => {
		const xs = [0, 1, 2, 3, 4];
		const ys = xs.map((x) => x * x - 2 * x + 1);
		const coeffs = fitPolynomial(xs, ys, 2);
		expect(coeffs[0]).toBeCloseTo(1, 3);
		expect(coeffs[1]).toBeCloseTo(-2, 3);
		expect(coeffs[2]).toBeCloseTo(1, 3);
	});

	it('returns an array of length order+1', () => {
		const xs = [0, 1, 2, 3, 4, 5];
		const ys = xs.map((x) => x * x);
		expect(fitPolynomial(xs, ys, 3)).toHaveLength(4);
	});
});

// ─────────────────────────────────────────────────────────────────────────────
// computeRSquared
// ─────────────────────────────────────────────────────────────────────────────

describe('computeRSquared', () => {
	it('returns 1 for a perfect fit', () => {
		const xs = [0, 1, 2, 3];
		const ys = [0, 1, 2, 3];
		const r2 = computeRSquared(xs, ys, (x) => x);
		expect(r2).toBeCloseTo(1, 6);
	});

	it('returns 0 for a constant prediction on non-constant data', () => {
		const xs = [0, 1, 2, 3];
		const ys = [0, 1, 4, 9];
		const meanY = ys.reduce((s, y) => s + y, 0) / ys.length;
		const r2 = computeRSquared(xs, ys, () => meanY);
		expect(r2).toBeCloseTo(0, 6);
	});

	it('returns 0 for empty arrays', () => {
		const r2 = computeRSquared([], [], (x) => x);
		expect(r2).toBe(0);
	});
});
