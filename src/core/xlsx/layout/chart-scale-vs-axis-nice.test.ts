import { describe, expect, it } from 'vitest';

import { niceValueAxisBounds } from '../../chart/axis-nice';
import { niceScale } from './chart-scale';

const inputs: [number, number][] = [];
for (const lo of [0, 1, 3, 10, 40, 85, -5, -90, -20]) {
	for (const hi of [1, 2.2, 7, 52, 100, 900, 1200, -1, 5]) {
		if (hi > lo) inputs.push([lo, hi]);
	}
}

describe('niceScale (xlsx) versus niceValueAxisBounds (chart)', () => {
	it('both anchor at zero for wide positive data and bracket the data', () => {
		for (const [lo, hi] of [
			[0, 52],
			[3, 900],
		] as const) {
			for (const intervals of [4, 5, 10]) {
				const a = niceScale(lo, hi, { maxIntervals: intervals });
				const b = niceValueAxisBounds(lo, hi, intervals);
				expect(a.min).toBe(0);
				expect(b.min).toBe(0);
				expect(a.max).toBeGreaterThanOrEqual(hi);
				expect(b.max).toBeGreaterThanOrEqual(hi);
			}
		}
	});

	it('the two intentionally differ, so the xlsx scale stays separate', () => {
		// Excel's candidate units are 1/2/5 x 10^n and the interval count is a ceiling
		// ("at most maxIntervals"); PowerPoint's include 2.5 and the count is a target
		// that the snap usually undershoots. xlsx tests pin the former.
		let differing = 0;
		for (const [lo, hi] of inputs) {
			for (const n of [4, 5, 10]) {
				const a = niceScale(lo, hi, { maxIntervals: n });
				const b = niceValueAxisBounds(lo, hi, n);
				if (a.min !== b.min || a.max !== b.max || a.majorUnit !== b.majorUnit) differing++;
			}
		}
		expect(differing).toBeGreaterThan(0);
	});

	it('shows the documented concrete difference: 2.5 steps are PowerPoint-only', () => {
		const xlsx = niceScale(0, 10, { maxIntervals: 5 });
		const chart = niceValueAxisBounds(0, 10, 5);
		expect(xlsx).toMatchObject({ max: 15, majorUnit: 5 });
		expect(chart).toMatchObject({ max: 12.5, majorUnit: 2.5 });
	});
});
