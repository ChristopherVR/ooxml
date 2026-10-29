// FORECAST.ETS family against values recorded from real Excel 16 on 17 series
// (__fixtures__/ets-series.json, answers in ets-excel.json; regenerate with ets-cases.ps1).
//
// Excel's optimiser is undocumented, so this engine reproduces the answers closely but not exactly.
// The tolerances below are the measured worst cases plus a small margin, not a promise of parity:
// they pin today's accuracy so a regression shows. Residual differences, honestly:
//  - forecasts: typically within 0.3% to 2% of Excel; up to 13% to 37% on short, irregular or
//    exponentially growing series, where Excel's parameters are not the least-squares optimum;
//  - FORECAST.ETS.CONFINT: within 0.3% to 40% (Excel's interval depends on its own parameters);
//  - FORECAST.ETS.SEASONALITY: 14 of 15 series agree; one noisy series reports a period where
//    Excel reports none;
//  - FORECAST.ETS.STAT: only the step size (type 8) and SMAPE (type 5) are compared; alpha, beta,
//    gamma and the error statistics come from this optimiser and differ from Excel's.
import { describe, expect, it } from 'vitest';
import { evaluateEts, type EtsKind, type EtsRow, recorded } from './ets-excel-compare';

const rows: EtsRow[] = recorded().flatMap(evaluateEts);
const only = (kind: EtsKind): EtsRow[] => rows.filter((r) => r.kind === kind);

/** Relative forecast tolerance per series; 0.03 unless listed. */
const FORECAST_TOLERANCE: Record<string, number> = {
	gaps: 0.14, // includes completion by zero fill, where the fitted model differs most
	duplicates: 0.25, // includes COUNT and SUM aggregation of uneven counts
	growth: 0.4, // exponential growth: Excel's trend is far more damped
	'short-6': 0.16,
	'step-change': 0.14,
	'blank-value': 0.17,
};
/** Relative confidence interval tolerance per series; 0.4 unless listed. */
const CONFINT_TOLERANCE: Record<string, number> = {
	growth: 0.8,
	'short-6': 0.5,
};

describe('FORECAST.ETS against Excel', () => {
	it('records 17 series', () => {
		expect(new Set(rows.map((r) => r.name)).size).toBe(17);
	});

	it('matches exact linear and constant series to rounding', () => {
		for (const r of rows.filter((x) => x.name === 'linear-exact' || x.name === 'constant')) {
			if (r.kind === 'forecast' || r.kind === 'confint') expect(r.error).toBeLessThan(1e-6);
		}
	});

	it.each(only('forecast').map((r) => [r.name, r.formula, r] as const))(
		'%s %s',
		(name, _formula, r) => {
			expect(r.error).toBeLessThanOrEqual(FORECAST_TOLERANCE[name] ?? 0.03);
		},
	);
});

describe('FORECAST.ETS.CONFINT against Excel', () => {
	it.each(only('confint').map((r) => [r.name, r.formula, r] as const))(
		'%s %s',
		(name, _formula, r) => {
			expect(r.error).toBeLessThanOrEqual(CONFINT_TOLERANCE[name] ?? 0.4);
		},
	);
});

describe('FORECAST.ETS.SEASONALITY against Excel', () => {
	it('agrees on 14 of 15 series; a noisy one reports a period where Excel reports none', () => {
		const wrong = only('seasonality').filter((r) => r.ours !== r.excel);
		expect(wrong.map((r) => r.name)).toEqual(['unsorted']);
	});
});

describe('FORECAST.ETS.STAT against Excel', () => {
	it('reports Excel step size', () => {
		for (const r of only('stat8')) expect(r.ours).toBe(r.excel);
	});

	it('is within 0.06 of Excel for SMAPE', () => {
		for (const r of only('stat5'))
			expect(Math.abs((r.ours as number) - (r.excel as number))).toBeLessThan(0.06);
	});

	it('keeps the smoothing parameters within Excel bounds', () => {
		for (const kind of ['stat1', 'stat2', 'stat3'] as const)
			for (const r of only(kind)) {
				expect(r.ours as number).toBeGreaterThanOrEqual(0);
				expect(r.ours as number).toBeLessThanOrEqual(1);
			}
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
