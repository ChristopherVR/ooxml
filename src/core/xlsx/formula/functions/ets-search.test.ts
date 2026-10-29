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

	it('moves beta to alpha - 0.001 only on 8 to 16 points', () => {
		const wantsBeta = bowl(0.5, 0.499);
		expect(chooseSmoothing(12, 0, wantsBeta).beta).toBeCloseTo(0.499, 9);
		expect(chooseSmoothing(30, 0, wantsBeta).beta).toBe(LOW);
	});

	it('falls back to the last value plus slope only when alpha lands on 0.998 on 8 to 16 points', () => {
		const top = (p: Smoothing): number => 1 + 1000 * (p.alpha - 0.999) ** 2 + (p.beta - LOW) ** 2;
		const short = chooseSmoothing(12, 0, top);
		expect([short.alpha, short.beta, short.drift]).toEqual([0.998, LOW, true]);
		// Beyond 16 points 0.998 is not taken: 0.9 stays.
		const long = chooseSmoothing(30, 0, top);
		expect([long.alpha, long.drift]).toEqual([0.9, false]);
	});

	it('keeps the earlier lattice for a season and for fewer than 8 points', () => {
		const seasonal = chooseSmoothing(24, 4, bowl(0.25));
		expect(seasonal.alpha).toBe(0.251);
		expect(seasonal.drift).toBe(false);
		expect(chooseSmoothing(6, 0, bowl(0.12)).alpha).toBe(0.126);
	});
});
