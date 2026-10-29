// Hard probes for the (alpha, beta) lattice of the FORECAST.ETS search (functions/ets-search-map.ts):
// 209 series (__fixtures__/ets-pairs-series.json, answers in ets-pairs-excel.json, written by
// scripts/xlsx/ets-pairs-series.ts) whose lowest hold-out cost is far lowest at one large-beta pair,
// and 231 more (ets-paths-*.json) that blend a series Excel forecasts from the pair with one it
// forecasts from the last value plus the slope, to cross the boundary between the two.
// The pair itself is found on every series (alpha 209 of 209 and every error statistic exact); which
// of the two forecasts Excel makes is only partly predicted (the floors are measured).
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { type EtsRecorded, evaluateEts } from './ets-excel-compare';

const load = (name: string): EtsRecorded[] =>
	JSON.parse(
		readFileSync(fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url)), 'utf8'),
	) as EtsRecorded[];

const tally = (set: EtsRecorded[]) => {
	const rows = set.flatMap(evaluateEts);
	const exact = (kind: RegExp) => rows.filter((r) => kind.test(r.kind) && r.error < 1e-9).length;
	const all = (kind: RegExp) => rows.filter((r) => kind.test(r.kind)).length;
	return {
		alpha: exact(/^stat1$/),
		beta: exact(/^stat2$/),
		stats: [exact(/^stat[4-7]$/), all(/^stat[4-7]$/)],
		forecasts: [exact(/^forecast$/), all(/^forecast$/)],
	};
};

describe('large-beta pair probes', () => {
	const pairs = load('ets-pairs-excel.json');
	const t = tally(pairs);

	it('records 209 series', () => {
		expect(pairs.length).toBe(209);
	});

	it('finds Excel alpha and every error statistic on all of them', () => {
		expect(t.alpha).toBe(209);
		expect(t.stats[0]).toBe(t.stats[1]);
	});

	it('reports Excel beta and forecasts the same on three in four', () => {
		expect(t.beta).toBeGreaterThanOrEqual(159);
		expect(t.forecasts[0]).toBeGreaterThanOrEqual(636);
	});
});

describe('series of 8 to 10 points (no hold-out)', () => {
	const short = load('ets-short-excel.json');
	const t = tally(short);

	it('finds Excel alpha, beta, every error statistic and every forecast on all 120 series', () => {
		expect(short.length).toBe(120);
		expect(t.alpha).toBe(120);
		expect(t.beta).toBe(120);
		expect(t.stats[0]).toBe(t.stats[1]);
		expect(t.forecasts[0]).toBe(t.forecasts[1]);
	});
});

describe('probes across the drift boundary', () => {
	const paths = load('ets-paths-excel.json');

	it('records 231 series', () => {
		expect(paths.length).toBe(231);
	});

	it('finds Excel alpha and every error statistic on all of them', () => {
		expect(tally(paths).alpha).toBe(231);
	});
});
