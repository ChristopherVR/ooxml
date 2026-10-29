import { describe, expect, it } from 'vitest';
import type { Smoothing } from './ets-core';
import { chooseSmoothing, LOW } from './ets-search';

/** A cost with its minimum at `alpha` (and the lowest beta), growing with the distance from it. */
const bowl =
	(alpha: number, beta = LOW) =>
	(p: Smoothing): number =>
		1 + (p.alpha - alpha) ** 2 + (p.beta - beta) ** 2;

describe('chooseSmoothing', () => {
	it('picks 0.9 when every parameter costs the same (an exact fit)', () => {
		for (const n of [5, 12, 30]) {
			const c = chooseSmoothing(n, 0, () => 3);
			expect([c.alpha, c.beta, c.drift]).toEqual([0.9, LOW, false]);
		}
	});

	it('lands on the lattice point nearest the minimum, never between', () => {
		expect(chooseSmoothing(30, 0, bowl(0.27)).alpha).toBe(0.25);
		expect(chooseSmoothing(30, 0, bowl(0.52)).alpha).toBe(0.5);
		expect(chooseSmoothing(30, 0, bowl(0.01)).alpha).toBe(0.002);
		expect(chooseSmoothing(12, 0, bowl(0.12)).alpha).toBe(0.126);
	});

	it('moves beta to alpha - 0.001 only where the cost falls a lot', () => {
		const wantsBeta = (p: Smoothing): number =>
			1 + 100 * ((p.alpha - 0.5) ** 2 + (p.beta - 0.499) ** 2);
		expect(chooseSmoothing(12, 0, wantsBeta).beta).toBeCloseTo(0.499, 9);
		expect(chooseSmoothing(30, 0, wantsBeta).beta).toBeCloseTo(0.499, 9);
		// A small gain is not enough: Excel keeps the lowest beta.
		const slight = (p: Smoothing): number => 1000 + (p.beta - LOW) ** 2 - 20 * p.beta * p.alpha;
		expect(chooseSmoothing(30, 0, slight).beta).toBe(LOW);
	});

	it('takes alpha 0.998 only when it clearly beats 0.9, and then forecasts from the last value', () => {
		const slightly = (p: Smoothing): number => 1 + 0.01 * (1 - p.alpha) ** 2;
		expect(chooseSmoothing(12, 0, slightly).alpha).toBe(0.9);
		const clearly = (p: Smoothing): number =>
			1 + 10 * (1 - p.alpha) ** 2 + 1000 * (p.beta - LOW) ** 2;
		const top = chooseSmoothing(12, 0, clearly);
		expect([top.alpha, top.beta, top.drift]).toEqual([0.998, LOW, true]);
		expect(chooseSmoothing(30, 0, clearly).drift).toBe(true);
	});

	it('keeps the earlier lattice for a season and for fewer than 8 points', () => {
		const seasonal = chooseSmoothing(24, 4, bowl(0.25));
		expect(seasonal.alpha).toBe(0.251);
		expect(seasonal.drift).toBe(false);
		expect(chooseSmoothing(6, 0, bowl(0.12)).alpha).toBe(0.126);
	});
});
