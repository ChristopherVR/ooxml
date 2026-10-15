import { describe, expect, it } from 'vitest';
import { solveGoal } from './goal-solver';

describe('solveGoal', () => {
	it('solves a linear equation', () => {
		const r = solveGoal((x) => 3 * x + 2, 17, 0);
		expect(r.found).toBe(true);
		expect(r.x).toBeCloseTo(5, 3);
	});

	it('is already done when the start meets the goal', () => {
		const r = solveGoal((x) => x * 2, 10, 5);
		expect(r).toMatchObject({ found: true, x: 5, iterations: 0 });
	});

	it('solves a quadratic from either side of the roots', () => {
		const up = solveGoal((x) => x * x, 16, 1);
		expect(up.found).toBe(true);
		expect(Math.abs(up.x)).toBeCloseTo(4, 2);
		const down = solveGoal((x) => x * x - 3 * x + 2, 0, 10);
		expect(down.found).toBe(true);
		expect([1, 2].some((root) => Math.abs(down.x - root) < 0.01)).toBe(true);
	});

	it('solves a loan payment for the amount borrowed', () => {
		const rate = 0.06 / 12;
		const pmt = (amount: number): number => (amount * rate) / (1 - Math.pow(1 + rate, -360));
		const r = solveGoal(pmt, 1500, 100000);
		expect(r.found).toBe(true);
		expect(r.x).toBeCloseTo(250187.1, -1);
	});

	it('reports no solution when the goal is unreachable', () => {
		const r = solveGoal((x) => x * x + 1, 0, 3);
		expect(r.found).toBe(false);
		expect(r.value).toBeGreaterThanOrEqual(1);
	});

	it('survives a function that errors somewhere', () => {
		const r = solveGoal((x) => (x < 0 ? Number.NaN : Math.sqrt(x)), 3, 1);
		expect(r.found).toBe(true);
		expect(r.x).toBeCloseTo(9, 2);
	});

	it('stops at the iteration limit', () => {
		const r = solveGoal((x) => Math.sin(x), 2, 0, { maxIterations: 5 });
		expect(r.found).toBe(false);
		expect(r.iterations).toBeLessThanOrEqual(5);
	});
});
