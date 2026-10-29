// How well the choice of alpha, beta and gamma (functions/ets-search.ts) reproduces Excel 16 on 491
// non-seasonal probe series (__fixtures__/ets-search-series.json, answers in ets-search-excel.json;
// regenerate with `bun scripts/xlsx/ets-search-series.ts` and ets-search-cases.ps1).
//
// What the probes established (docs/xlsx-parity-status.md has the evidence): the objective is the
// one-step squared error from the least-squares trend; Excel lands on a coarse lattice and stops
// early on many series, so these are pinned accuracy floors, not a promise of parity. Before this
// search: 212 of 491 alphas, 832 of 1964 forecasts, median forecast error 9.5e-4, worst 0.205.
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

	it('finds Excel alpha on half of the series', () => {
		expect(exact(alphas)).toBeGreaterThanOrEqual(240);
	});

	it('finds it on three quarters of the series of 8 to 16 points', () => {
		const short = alphas.filter((r) => (lengthOf.get(r.name) as number) <= 16);
		expect(short.length).toBe(153);
		expect(exact(short)).toBeGreaterThanOrEqual(110);
	});

	it('keeps beta at the lowest value on nine series in ten, as Excel does', () => {
		const betas = rows.filter((r) => r.kind === 'stat2');
		expect(exact(betas)).toBeGreaterThanOrEqual(438);
	});

	it('forecasts the same as Excel on half of the calls and within 0.12 on all', () => {
		expect(exact(forecasts)).toBeGreaterThanOrEqual(940);
		expect(Math.max(...forecasts.map((r) => r.error))).toBeLessThan(0.12);
		expect(median(forecasts)).toBeLessThan(3e-4);
	});

	it('reproduces the error statistics wherever it found the same parameters', () => {
		const same = new Set(alphas.filter((r) => r.error < 1e-9).map((r) => r.name));
		const sameBeta = new Set(
			rows.filter((r) => r.kind === 'stat2' && r.error < 1e-9).map((r) => r.name),
		);
		const both = [...same].filter((n) => sameBeta.has(n));
		expect(both.length).toBeGreaterThan(200);
		// Excel makes about a tenth of the forecasts from the last value and the plain slope (alpha 1, beta 0)
		// and measures their errors some other way: a few of them have the parameters it reports.
		const off = rows.filter((x) => x.kind === 'stat7' && both.includes(x.name) && x.error > 1e-6);
		expect(off.length).toBeLessThanOrEqual(8);
	});
});
