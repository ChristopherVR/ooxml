// What the FORECAST.ETS parameter search minimises, checked against probes recorded from Excel 16:
// the training one-step error plus the hold-out forecast error (functions/ets-validate.ts).
// Series are written by scripts/xlsx/ets-{morph,spike,scale}-series.ts and recorded with
// __fixtures__/ets-search-cases.ps1. The floors are measured accuracy, not a promise of parity.
import { describe, expect, it } from 'vitest';
import { type EtsRecorded, probeRecorded } from './ets-excel-compare';
import { etsRun, fitEts } from './functions/ets-core';
import { etsValidationCost } from './functions/ets-validate';

const excel = (c: EtsRecorded, statistic: number): number =>
	c.results.find((r) => r.formula.startsWith(`FORECAST.ETS.STAT({V},{L},${statistic},`))
		?.value as number;

const values = (c: EtsRecorded): number[] => c.values as number[];
const sameAlpha = (c: EtsRecorded): boolean =>
	Math.abs(fitEts(values(c), 0).alpha - excel(c, 1)) < 1e-9;
const sameBeta = (c: EtsRecorded): boolean =>
	Math.abs(fitEts(values(c), 0).beta - excel(c, 2)) < 1e-9;

describe('etsValidationCost', () => {
	const init = { level: 0, trend: 1, season: [] };
	const p = { alpha: 0.5, beta: 0, gamma: 0 };
	const y = [1, 2, 3, 5, 9];

	it('is the one-step squared error without a hold-out', () => {
		const run = etsRun(y, 0, init, p);
		expect(etsValidationCost(y, 0, init, p, 0)).toBeCloseTo(run.sse, 12);
		expect(run.sse).toBeCloseTo(13.25, 12);
	});

	it('forecasts the hold-out ahead from the end of the training part without updating', () => {
		// Training 1, 2, 3 is fitted exactly (level 3, trend 1): forecasts 4 and 5 miss 5 and 9.
		expect(etsValidationCost(y, 0, init, p, 2)).toBeCloseTo(1 + 16, 12);
	});

	it('adds the season of the forecast position', () => {
		const seasonal = { level: 0, trend: 0, season: [1, 2] };
		// Fitted 1, 2, 1, 2 exactly with gamma 0; hold-out of 2 forecasts 1 and 2.
		expect(etsValidationCost([1, 2, 1, 2, 1, 5], 2, seasonal, p, 2)).toBeCloseTo(0 + 9, 12);
	});
});

describe('FORECAST.ETS parameter search on series morphing from a random walk to noise', () => {
	const morph = probeRecorded('morph');

	it('records 484 series', () => {
		expect(morph.length).toBe(484);
	});

	it('finds Excel alpha on every series', () => {
		// 51 of the 63 misses were one family (Excel's 1/6 on a bimodal cost surface of 40 points), which
		// the lattice and the report map (functions/ets-search-map.ts) now reproduce.
		expect(morph.filter(sameAlpha).length).toBeGreaterThanOrEqual(484);
	});

	it('finds Excel beta on every series', () => {
		expect(morph.filter(sameBeta).length).toBeGreaterThanOrEqual(484);
	});
});

describe('FORECAST.ETS parameter search on a constant with one spike or step', () => {
	const spikes = probeRecorded('spike');

	it('finds Excel alpha on all but two of 136 series', () => {
		// The cost sees the hold-out only as forecasts: a spike or step among the last points leaves
		// the training part constant and Excel's alpha falls to 0.1, 0.126 and 0.25 as it nears the end.
		expect(spikes.length).toBe(136);
		expect(spikes.filter(sameAlpha).length).toBeGreaterThanOrEqual(134);
	});

	it('reproduces the late spikes of 30 points', () => {
		const late = spikes.filter((c) => /^spike-30-(2[4-9])$/.test(c.name));
		expect(late.map((c) => excel(c, 1))).toEqual([0.1, 0.1, 0.1, 0.126, 0.126, 0.25]);
		expect(late.every(sameAlpha)).toBe(true);
	});
});

describe('FORECAST.ETS parameter search is invariant to scale and sign, as Excel is', () => {
	const scaled = probeRecorded('scale');
	const groups = new Map<string, EtsRecorded[]>();
	for (const c of scaled) {
		const [base, kind] = c.name.split('|') as [string, string];
		if (kind === 'rev') continue;
		groups.set(base, [...(groups.get(base) ?? []), c]);
	}

	it('records 49 base series in 7 versions each (scales 1e-4 to 1e7, and negated)', () => {
		expect(groups.size).toBe(49);
		for (const g of groups.values()) expect(g.length).toBe(7);
	});

	it('gives Excel the same alpha and beta for every version of a series', () => {
		for (const g of groups.values())
			expect(new Set(g.map((c) => `${excel(c, 1)}/${excel(c, 2)}`)).size).toBe(1);
	});

	it('gives the same alpha and beta for every version of a series', () => {
		for (const g of groups.values()) {
			const fits = g.map((c) => fitEts(values(c), 0));
			expect(new Set(fits.map((f) => `${f.alpha}/${f.beta}`)).size).toBe(1);
		}
	});

	it('finds Excel alpha on 99% of the series, reversed ones included', () => {
		expect(scaled.length).toBe(392);
		expect(scaled.filter(sameAlpha).length).toBeGreaterThanOrEqual(388);
	});
});
