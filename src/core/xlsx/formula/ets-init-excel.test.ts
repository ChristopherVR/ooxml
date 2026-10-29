// Starting states and seasonality detection of FORECAST.ETS against what real Excel 16 did on the
// probes in __fixtures__/ets-init-series.json (answers in ets-init-excel.json; regenerate with
// ets-init-cases.ps1). Excel's starting state is recovered from its forecasts (see
// ets-init-compare.ts), so these are comparisons with Excel's internals, not only its outputs.
//
// Established: the filter form (the recovered state reproduces the first observation exactly on
// every consistent series), the OLS starting slope for 3 and 7 or more points, the 6-point
// identity in `sixPointTrend`, and that Excel's seasonal starting state is a fixed linear function
// of the data that does not depend on the fitted parameters (so it is not a least-squares fit of
// the filter). Not established: that function for 4 and 5 points and for seasonal series.
import { describe, expect, it } from 'vitest';
import { recordedDetections, recoveredInits } from './ets-init-compare';
import { etsRun } from './functions/ets-core';
import { initialStates } from './functions/ets-init';
import { detectSeasonality } from './functions/ets-season';

const cases = recoveredInits();
const consistent = cases.filter((c) => Math.abs(c.firstError) < 1e-6);

/** Our starting state for the case, as level plus season per position and the trend. */
function ours(c: (typeof cases)[number]): { c: number[]; trend: number } {
	const s = initialStates(c.y, c.m, c, etsRun);
	return {
		c: c.m >= 2 ? s.season.map((x) => x + s.level) : [s.level],
		trend: s.trend,
	};
}
const trendError = (name: RegExp): number[] =>
	consistent.filter((c) => name.test(c.name)).map((c) => Math.abs(ours(c).trend - c.trend));

describe('FORECAST.ETS starting states against Excel', () => {
	it('records 184 probes of which 172 give a consistent state', () => {
		expect(cases.length).toBe(184);
		// The other 12 are series fitted with a large alpha plus gamma, where recovering the start
		// from the end amplifies rounding until the first-value law no longer holds.
		expect(consistent.length).toBe(172);
	});

	it('uses the least-squares slope for 3, 7 and 8 points and the 6-point identity for 6', () => {
		for (const n of [3, 6, 7, 8])
			for (const e of trendError(new RegExp(`^n${n}-`))) expect(e).toBeLessThan(1e-8);
	});

	it('keeps the measured distance from Excel for 4 and 5 points', () => {
		expect(Math.max(...trendError(/^n4-/))).toBeLessThan(0.14);
		expect(Math.max(...trendError(/^n5-/))).toBeLessThan(0.32);
	});

	it('keeps the measured distance from Excel for seasonal series', () => {
		const seasonal = consistent.filter((c) => c.m >= 2 && c.gamma > 1e-9);
		const meanTrend = seasonal.reduce((s, c) => s + Math.abs(ours(c).trend - c.trend), 0);
		expect(seasonal.length).toBe(107);
		expect(meanTrend / seasonal.length).toBeLessThan(0.22);
	});
});

describe('FORECAST.ETS.SEASONALITY against Excel on designed series', () => {
	const detections = recordedDetections();
	const wrong = detections.filter((d) => detectSeasonality(d.y) !== d.excel);

	it('reproduces every noiseless wave, trend and impulse train', () => {
		const clean = detections.filter((d) => !/noise|noisy|det-two/.test(d.name));
		const missed = clean.filter((d) => detectSeasonality(d.y) !== d.excel);
		expect(missed.map((d) => d.name)).toEqual([]);
	});

	it('agrees on 100 of the 111 probes', () => {
		expect(detections.length - wrong.length).toBeGreaterThanOrEqual(100);
	});

	it('does not reproduce the periods Excel reports on short noisy series', () => {
		// Excel finds a period on a third of the white-noise series of 8 to 36 points (and none of
		// 48 or more); this detector reports none on most of them. Pinned so a change shows.
		const noise = detections.filter((d) => d.name.startsWith('det-noise-'));
		const found = noise.filter((d) => d.excel !== 0);
		expect(found.length).toBe(10);
		expect(found.filter((d) => detectSeasonality(d.y) === d.excel).length).toBeLessThanOrEqual(4);
	});
});
