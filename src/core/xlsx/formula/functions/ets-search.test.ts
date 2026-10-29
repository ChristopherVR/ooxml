import { describe, expect, it } from 'vitest';
import type { Smoothing } from './ets-core';
import { chooseSmoothing, LOW } from './ets-search';
import { CANDIDATES, reportedFor } from './ets-search-map';

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

	it('reports 0.1 and 0.099 for 3 points, whatever the values', () => {
		const c = chooseSmoothing(3, 0, bowl(0.5));
		expect([c.alpha, c.beta, c.drift]).toEqual([0.1, 0.099, false]);
	});

	it('lands on the lattice point nearest the minimum, never between', () => {
		expect(chooseSmoothing(30, 0, bowl(0.27)).alpha).toBe(0.25);
		expect(chooseSmoothing(30, 0, bowl(0.52)).alpha).toBe(0.5);
		expect(chooseSmoothing(30, 0, bowl(0.01)).alpha).toBe(0.002);
		expect(chooseSmoothing(12, 0, bowl(0.12)).alpha).toBe(0.126);
		expect(chooseSmoothing(6, 0, bowl(0.12)).alpha).toBe(0.126);
	});

	it('reports a large beta next to alpha as 0.001 and forecasts the last value plus the slope', () => {
		const c = chooseSmoothing(30, 0, bowl(0.25, 0.249));
		expect([c.alpha, c.beta, c.drift]).toEqual([0.25, LOW, true]);
		// The error statistics still use the pair.
		expect(c.fit).toEqual({ alpha: 0.25, beta: 0.249, gamma: LOW });
	});

	it('reports the larger betas of alpha 0.998 as they are', () => {
		const c = chooseSmoothing(30, 0, bowl(0.998, 0.499));
		expect([c.alpha, c.beta, c.drift]).toEqual([0.998, 0.499, false]);
	});

	it('maps the optimum 0.25 with 0.099 to the 1/6 Excel reports', () => {
		const c = chooseSmoothing(30, 0, bowl(0.25, 0.125));
		expect([c.alpha, c.beta, c.drift]).toEqual([0.166667, LOW, true]);
	});

	it('takes alpha 0.998 with the lowest beta only when it clearly beats 0.9', () => {
		const slightly = (p: Smoothing): number =>
			1 + 0.01 * (1 - p.alpha) ** 2 + 1000 * (p.beta - LOW) ** 2;
		expect(chooseSmoothing(12, 0, slightly).alpha).toBe(0.9);
		const clearly = (p: Smoothing): number =>
			1 + 10 * (1 - p.alpha) ** 2 + 1000 * (p.beta - LOW) ** 2;
		expect(chooseSmoothing(12, 0, clearly).alpha).toBe(0.998);
	});

	it('starts the lattice with the eight alphas at the lowest beta', () => {
		expect(CANDIDATES.slice(0, 8).map(([a]) => a)).toEqual([
			0.002, 0.1, 0.126, 0.25, 0.5, 0.75, 0.9, 0.998,
		]);
		expect(CANDIDATES.every(([a, b]) => b < a)).toBe(true);
	});

	it('keeps the earlier lattice for a season', () => {
		const seasonal = chooseSmoothing(24, 4, bowl(0.25));
		expect(seasonal.alpha).toBe(0.251);
		expect(seasonal.drift).toBe(false);
	});
});

describe('reportedFor', () => {
	it('reports the pairs of the base lattice as they are', () => {
		for (const a of [0.002, 0.1, 0.126, 0.25, 0.5, 0.75, 0.9]) {
			const out = reportedFor(a, LOW, 30);
			expect([out.alpha, out.beta, out.drift]).toEqual([a, LOW, false]);
		}
	});

	it('splits the optimum 0.75 with 0.749 by length, as the three recorded series do', () => {
		expect(reportedFor(0.75, 0.749, 13).drift).toBe(false);
		expect(reportedFor(0.75, 0.749, 24).drift).toBe(true);
	});
});
