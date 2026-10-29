// Starting states and seasonality detection of FORECAST.ETS against what real Excel 16 did on the
// probes in __fixtures__/ets-init-series.json (answers in ets-init-excel.json; regenerate with
// ets-init-cases.ps1). Excel's starting state is recovered from its forecasts (see
// ets-init-compare.ts), so these are comparisons with Excel's internals, not only its outputs.
//
// Established: the filter form (the recovered state reproduces the first observation exactly on
// every consistent series), the starting slope of every series without a season (`shortSeriesTrend`:
// least squares from 7 values, the line through 7 values extended from the data for 4 to 6), and
// that Excel's seasonal starting state is a fixed linear function of the data that does not depend
// on the fitted parameters (so it is not a least-squares fit of the filter). Not established: that
// function for seasonal series; ets-init-maps.test.ts records the laws found so far.
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

	it('reproduces the starting trend for 3 to 8 values without a season', () => {
		for (const n of [3, 4, 5, 6, 7, 8])
			for (const e of trendError(new RegExp(`^n${n}-`))) expect(e).toBeLessThan(1e-8);
	});

	it('keeps the measured distance from Excel for seasonal series', () => {
		const seasonal = consistent.filter((c) => c.m >= 2 && c.gamma > 1e-9);
		const meanTrend = seasonal.reduce((s, c) => s + Math.abs(ours(c).trend - c.trend), 0);
		expect(seasonal.length).toBe(107);
		expect(meanTrend / seasonal.length).toBeLessThan(0.14);
	});
});

describe('FORECAST.ETS on 5 values with the parameters Excel reports', () => {
	// Excel (ets-excel.json, "short-5") reports alpha 0.1, beta 0.001 and forecasts 5.2258, 6.0416 and
	// 8.4890 at the next three points and an interval of 1.8509 around the first: with the exact
	// starting trend the filter alone reproduces them (the parameter search is a separate matter).
	it('gives Excel forecasts, RMSE and interval to the last digits', () => {
		const y = [1, 3, 2, 5, 4];
		const trend = initialStates(y, 0, { alpha: 0.1, beta: 0.001, gamma: 0 }, etsRun).trend;
		const run = etsRun(
			y,
			0,
			{ level: (y[0] as number) - trend, trend, season: [] },
			{ alpha: 0.1, beta: 0.001, gamma: 0 },
		);
		const forecast = (h: number): number => run.level + h * run.trend;
		expect(forecast(1)).toBeCloseTo(5.225847882010245, 11);
		expect(forecast(2)).toBeCloseTo(6.041627131381917, 11);
		expect(forecast(5)).toBeCloseTo(8.488964879496933, 11);
		expect(Math.sqrt(run.sse / y.length)).toBeCloseTo(0.9443512822685144, 11);
		expect(1.959963984540054 * Math.sqrt(run.sse / y.length)).toBeCloseTo(1.8508945020005063, 9);
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
