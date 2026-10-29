// FORECAST.ETS family against values recorded from real Excel 16 on 34 series
// (__fixtures__/ets-series.json, answers in ets-excel.json; regenerate with ets-cases.ps1).
//
// What is identified (and reproduced to the last digit on the series listed in EXACT): the filter
// (level += trend + alpha*e, trend += beta*e, season += gamma*e), the first fitted value being the
// first observation, the trend starting at the least-squares slope of the series (for seasons: of the series without the exact seasonal start of ets-init-chain.ts), the error
// statistics being measured on the last points only (see `etsHoldout`) and the interval
// being RMSE * sqrt(1 + sum of (alpha + j*beta)^2). What is not: Excel's search for alpha, beta and
// gamma stops early on some series, and its starting values for seasonal models outside ets-init-chain.ts and ets-init-three.ts (seasons of 2 or 3, say) and for fewer than
// 7 points are not the least-squares ones (ets-init-excel.test.ts recovers them from Excel; the search
// is measured on 1100 more series in forecast-ets-search.test.ts and forecast-ets-validate.test.ts). The tolerances below are the measured worst cases plus a
// small margin, not a promise of parity: they pin today's accuracy so a regression shows.
import { describe, expect, it } from 'vitest';
import { evaluateEts, type EtsKind, type EtsRow, recorded } from './ets-excel-compare';
import { etsHoldout } from './functions/ets-core';

const rows: EtsRow[] = recorded().flatMap(evaluateEts);
const only = (kind: EtsKind): EtsRow[] => rows.filter((r) => r.kind === kind);

/** Series on which every statistic, forecast and interval equals Excel's. */
const EXACT = [
	'linear-exact',
	'constant',
	'growth',
	'step-0-10',
	'walk-7',
	'trend-9',
	'walk-12',
	'trend-24',
	'walk-100',
	'unsorted',
	'random-walk',
	'spike',
	'step-change',
	'short-5',
	'walk-20-anomaly',
	'walk-50',
];

/** Relative forecast tolerance per series (the measured worst case plus 3%); 1e-9 unless listed. */
const FORECAST_TOLERANCE: Record<string, number> = {
	'flat-noisy': 0.0387, // Excel finds a period of 2 in this noise, which no detector here reproduces
	gaps: 0.145, // zero-filled gaps: Excel finds no period, the detector here finds 5 (see ets-season.ts)
	duplicates: 0.0484, // the calls that detect a season (those with seasonality 0 are exact)
	'short-6': 0.0249, // the start is exact; Excel ends at alpha 0.998 with beta 0.099, the cost says 0.002
	'yearly-step': 0.055, // Excel reports the pair 0.9 and 0.899 and forecasts from it; this reports 0.001
	'seasonal3-noisy': 0.0145,
	'walk-13': 0.124, // Excel reports the pair 0.75 and 0.749 here; the pair drifts in most other series
	'season2-9': 0.0037,
	'season3-9': 0.0079,
	'season3-11': 0.0111,
};
/** Relative confidence interval tolerance per series (the measured worst case plus 3%); 1e-9 unless listed. */
const CONFINT_TOLERANCE: Record<string, number> = {
	'flat-noisy': 0.201,
	gaps: 0.276, // zero completion: Excel finds no period on the gap-filled grid, this finds 5
	'short-6': 0.176,
	'yearly-step': 0.539,
	'walk-13': 0.509,
	'season2-9': 0.068,
	'season3-9': 0.073,
	'season3-11': 0.131,
};

describe('FORECAST.ETS against Excel', () => {
	it('records 34 series', () => {
		expect(new Set(rows.map((r) => r.name)).size).toBe(34);
	});

	it('reproduces the identified model to the last digit', () => {
		for (const r of rows.filter((x) => EXACT.includes(x.name)))
			if (r.kind !== 'seasonality' && r.kind !== 'stat8') expect(r.error).toBeLessThan(1e-6);
	});

	it.each(only('forecast').map((r) => [r.name, r.formula, r] as const))(
		'%s %s',
		(name, _formula, r) => {
			expect(r.error).toBeLessThanOrEqual(FORECAST_TOLERANCE[name] ?? 1e-9);
		},
	);
});

describe('FORECAST.ETS.CONFINT against Excel', () => {
	it.each(only('confint').map((r) => [r.name, r.formula, r] as const))(
		'%s %s',
		(name, _formula, r) => {
			expect(r.error).toBeLessThanOrEqual(CONFINT_TOLERANCE[name] ?? 1e-9);
		},
	);
});

describe('FORECAST.ETS.SEASONALITY against Excel', () => {
	it('agrees on 13 of 15 series; Excel reports a period on two noisy ones where this detector reports none', () => {
		const wrong = only('seasonality').filter((r) => r.ours !== r.excel);
		expect(wrong.map((r) => r.name)).toEqual(['flat-noisy', 'gaps']);
	});
});

describe('FORECAST.ETS.STAT against Excel', () => {
	it('reports Excel step size', () => {
		for (const r of only('stat8')) expect(r.ours).toBe(r.excel);
	});

	it('is within 0.06 of Excel for SMAPE (other than on the anomalous walk)', () => {
		for (const r of only('stat5').filter((x) => x.name !== 'walk-20-anomaly'))
			expect(Math.abs((r.ours as number) - (r.excel as number))).toBeLessThan(0.06);
	});

	it('keeps the smoothing parameters within Excel bounds', () => {
		for (const kind of ['stat1', 'stat2', 'stat3'] as const)
			for (const r of only(kind)) {
				expect(r.ours as number).toBeGreaterThanOrEqual(0);
				expect(r.ours as number).toBeLessThanOrEqual(1);
			}
	});

	it('finds Excel alpha on most series', () => {
		const alpha = only('stat1');
		const same = alpha.filter((r) => Math.abs((r.ours as number) - (r.excel as number)) < 1e-9);
		expect(same.length).toBeGreaterThanOrEqual(15);
	});

	it('measures the error statistics on the last points only, as Excel does', () => {
		expect([10, 11, 20, 30, 60, 100].map((n) => etsHoldout(n, 0))).toEqual([0, 4, 8, 10, 12, 20]);
		// Two seasons at least: 35 points with a season of 7 hold out 14.
		expect(etsHoldout(35, 7)).toBe(14);
	});
});

describe('FORECAST.ETS argument handling matched to Excel', () => {
	it('treats an empty value as a missing point and completes it', () => {
		const row = rows.find(
			(r) => r.name === 'blank-value' && r.formula.startsWith('FORECAST.ETS(9,{V},{L})'),
		);
		expect(row?.ours).toBeCloseTo(19, 6);
	});

	it('aggregates duplicate times as Excel numbers them (5 is MEDIAN, 6 is MIN), exactly', () => {
		const model = rows.filter(
			(r) => r.name === 'duplicates' && /^FORECAST\.ETS\(13,\{V\},\{L\},0,1,\d\)$/.test(r.formula),
		);
		expect(model.length).toBe(7);
		for (const r of model) expect(r.error).toBeLessThan(1e-9);
	});

	it('steps a monthly date timeline by calendar months', () => {
		const row = rows.find((r) => r.name === 'monthly-dates-seasonal12' && r.kind === 'stat8');
		expect(row?.ours).toBe(31);
		expect(
			rows.filter((r) => r.name === 'monthly-dates-seasonal12' && r.kind === 'forecast').length,
		).toBe(2);
	});
});
