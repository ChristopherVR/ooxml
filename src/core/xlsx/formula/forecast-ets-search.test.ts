// How well the choice of alpha, beta and gamma (functions/ets-search.ts) reproduces Excel 16 on 491
// non-seasonal probe series (__fixtures__/ets-search-series.json, answers in ets-search-excel.json;
// regenerate with `bun scripts/xlsx/ets-search-series.ts` and ets-search-cases.ps1).
//
// What the probes established (docs/xlsx-parity-status.md has the evidence): the cost is the
// training one-step squared error plus the hold-out forecast error from the least-squares trend
// (functions/ets-validate.ts), minimised over a fixed lattice of (alpha, beta) pairs, and what Excel
// reports is a function of the optimum (functions/ets-search-map.ts). These are pinned accuracy
// floors, not a promise of parity. Before the lattice and the report map: 471 alphas, 443 betas,
// 1716 of 1964 forecasts and a worst forecast error of 0.107. With the one-step error of the whole
// series instead: 243 of 491 alphas, 948 of 1964 forecasts, worst forecast error 0.103.
import { describe, expect, it } from 'vitest';
import { evaluateEts, type EtsRow, searchRecorded } from './ets-excel-compare';

const series = searchRecorded();
const rows: EtsRow[] = series.flatMap(evaluateEts);
const lengthOf = new Map(series.map((c) => [c.name, c.values.length]));
const alphas = rows.filter((r) => r.kind === 'stat1');
const forecasts = rows.filter((r) => r.kind === 'forecast');
const exact = (rs: EtsRow[]): number => rs.filter((r) => r.error < 1e-9).length;
const median = (rs: EtsRow[]): number =>
	rs.map((r) => r.error).sort((a, b) => a - b)[rs.length >> 1] as number;

describe('FORECAST.ETS parameter search against Excel', () => {
	it('records 491 series', () => {
		expect(series.length).toBe(491);
	});

	it('finds Excel alpha on 98% of the series', () => {
		expect(exact(alphas)).toBeGreaterThanOrEqual(483);
	});

	it('finds it on 99% of the series of 8 to 16 points', () => {
		const short = alphas.filter((r) => (lengthOf.get(r.name) as number) <= 16);
		expect(short.length).toBe(153);
		expect(exact(short)).toBeGreaterThanOrEqual(151);
	});

	it('finds Excel beta on 96% of the series', () => {
		const betas = rows.filter((r) => r.kind === 'stat2');
		expect(exact(betas)).toBeGreaterThanOrEqual(471);
	});

	it('forecasts the same as Excel on 95% of the calls and within 0.11 on all', () => {
		expect(exact(forecasts)).toBeGreaterThanOrEqual(1868);
		expect(Math.max(...forecasts.map((r) => r.error))).toBeLessThan(0.11);
		expect(median(forecasts)).toBeLessThan(1e-9);
	});

	it('reproduces the error statistics wherever it found the same parameters', () => {
		const same = new Set(alphas.filter((r) => r.error < 1e-9).map((r) => r.name));
		const sameBeta = new Set(
			rows.filter((r) => r.kind === 'stat2' && r.error < 1e-9).map((r) => r.name),
		);
		const both = [...same].filter((n) => sameBeta.has(n));
		expect(both.length).toBeGreaterThanOrEqual(467);
		// Where the optimum has a large beta Excel forecasts from the last value and the plain slope but
		// measures the errors with the pair; 6 series report an alpha of 1/6 whose pair is not known.
		const off = rows.filter((x) => x.kind === 'stat7' && both.includes(x.name) && x.error > 1e-6);
		expect(off.length).toBeLessThanOrEqual(6);
	});
});
