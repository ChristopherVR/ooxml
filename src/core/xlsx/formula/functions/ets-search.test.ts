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
		const c = chooseSmoothing(30, 0, bowl(0.166667, 0.0833333));
		expect([c.alpha, c.beta, c.drift]).toEqual([0.166667, LOW, true]);
	});

	it('never fits alpha 0.998 with the lowest beta, however much lower its cost is', () => {
		const clearly = (p: Smoothing): number =>
			1 + 10 * (1 - p.alpha) ** 2 + 1000 * (p.beta - LOW) ** 2;
		expect(chooseSmoothing(12, 0, clearly).alpha).toBe(0.9);
	});

	it('starts the lattice with the seven alphas at the lowest beta', () => {
		expect(CANDIDATES.slice(0, 7).map(([a]) => a)).toEqual([
			0.002, 0.1, 0.126, 0.25, 0.5, 0.75, 0.9,
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
			const out = reportedFor(a, LOW, { shift: 0, t: 0 });
			expect([out.alpha, out.beta, out.drift]).toEqual([a, LOW, false]);
		}
	});

	it('drifts at the pairs next to alpha - 0.001 that mostly drift, and decides the rest by the shape', () => {
		const trending = { shift: 3, t: 30 };
		const flat = { shift: 0.5, t: 1 };
		expect(reportedFor(0.25, 0.249, trending).drift).toBe(true);
		expect(reportedFor(0.5, 0.25, trending).drift).toBe(false);
		expect(reportedFor(0.5, 0.25, flat).drift).toBe(true);
		expect(reportedFor(0.5, 0.25, flat).beta).toBe(LOW);
	});
});
