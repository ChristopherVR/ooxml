// FORECAST.ETS family against values recorded from real Excel 16 on 34 series
// (__fixtures__/ets-series.json, answers in ets-excel.json; regenerate with ets-cases.ps1).
//
// What is identified (and reproduced to the last digit on the series listed in EXACT): the filter
// (level += trend + alpha*e, trend += beta*e, season += gamma*e), the first fitted value being the
// first observation, the trend starting at the least-squares slope of the series, the error
// statistics being measured on the last points only (see `etsHoldout`) and the interval
// being RMSE * sqrt(1 + sum of (alpha + j*beta)^2). What is not: Excel's search for alpha, beta and
// gamma stops early on some series, and its starting values for a seasonal model and for fewer than
// 7 points are not the least-squares ones (ets-init-excel.test.ts recovers them from Excel). The tolerances below are the measured worst cases plus a
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
	'spike',
	'step-0-10',
	'walk-7',
	'trend-9',
	'walk-12',
	'trend-24',
	'walk-100',
];

/** Relative forecast tolerance per series; 0.003 unless listed. */
const FORECAST_TOLERANCE: Record<string, number> = {
	'flat-noisy': 0.04, // Excel finds a period of 2 in this noise, which no detector here reproduces
	'seasonal4-trend': 0.025, // seasonal starting values differ from Excel's
	'monthly-dates-seasonal12': 0.004,
	'seasonal7-daily': 0.006,
	gaps: 0.1, // zero-filled gaps: Excel picks a smaller alpha than the best fit
	duplicates: 0.06, // includes SUM aggregation of uneven counts, where Excel stops early
	'short-6': 0.025, // fewer than 7 points: Excel's starting trend is not the least-squares slope
	'random-walk': 0.012,
	'step-change': 0.012,
	'yearly-step': 0.012,
	'seasonal3-noisy': 0.02,
	'blank-value': 0.2, // zero completion of the missing point
	'short-5': 0.045,
	'walk-20-anomaly': 0.07, // Excel's search stops at alpha 0.25 here
	'walk-13': 0.008,
	'walk-50': 0.022,
	'season2-9': 0.008,
	'season3-9': 0.012,
	'season4-8': 0.01,
	'season5-10': 0.055,
	'season3-11': 0.015,
};
/** Relative confidence interval tolerance per series; 0.002 unless listed. */
const CONFINT_TOLERANCE: Record<string, number> = {
	'flat-noisy': 0.2,
	'seasonal4-trend': 0.11,
	'monthly-dates-seasonal12': 0.35,
	'seasonal7-daily': 0.2,
	gaps: 0.38, // zero completion: Excel's period 4 on the gap-filled grid is not reproduced
	'short-6': 0.18,
	'random-walk': 0.3,
	'step-change': 0.21,
	'yearly-step': 0.27,
	'seasonal3-noisy': 0.23,
	'short-5': 0.03,
	'walk-20-anomaly': 0.22,
	'walk-13': 0.12,
	'walk-25-anomaly': 0.15,
	'walk-50': 0.15,
	'season2-9': 0.08,
	'season3-9': 0.08,
	'season4-8': 0.26,
	'season5-10': 0.22,
	'season3-11': 0.13,
	'linear-noisy': 0.002,
	'walk-7': 0.002,
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
			expect(r.error).toBeLessThanOrEqual(FORECAST_TOLERANCE[name] ?? 0.003);
		},
	);
});

describe('FORECAST.ETS.CONFINT against Excel', () => {
	it.each(only('confint').map((r) => [r.name, r.formula, r] as const))(
		'%s %s',
		(name, _formula, r) => {
			expect(r.error).toBeLessThanOrEqual(CONFINT_TOLERANCE[name] ?? 0.002);
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

	it('steps a monthly date timeline by calendar months', () => {
		const row = rows.find((r) => r.name === 'monthly-dates-seasonal12' && r.kind === 'stat8');
		expect(row?.ours).toBe(31);
		expect(
			rows.filter((r) => r.name === 'monthly-dates-seasonal12' && r.kind === 'forecast').length,
		).toBe(2);
	});
});
