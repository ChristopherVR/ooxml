// Excel's FORECAST.ETS.SEASONALITY on 352 designed series (__fixtures__/ets-detect-series.json,
// answers in ets-detect-excel.json, written by scripts/xlsx/ets-detect-series.ts): white noise at 16
// lengths from 8 to 60 with 16 seeds each, and noisy sinusoids of period 2 to 12; and on 768
// transformed copies of the noise (ets-detect-variants-excel.json, scripts/xlsx/ets-detect-variants.ts).
//
// Measured about the rule, which is still not reproduced:
//   - the answer is invariant to negating, reversing, rescaling and shifting the series (768 of 768
//     copies equal the original), so it is a function of the shape that is symmetric in time:
//     autocorrelation or spectrum, not the fit of a seasonal ETS model (which is not symmetric under
//     reversal);
//   - on white noise the reported period is the rounding of n / k (halves rounded down) for the
//     Fourier frequency k with the largest periodogram ordinate of the linearly detrended series on
//     most series where one is reported (the autocorrelation maximum does as well only on half);
//   - whether a period is reported at all is not a threshold on that ordinate's share of the total
//     (it separates the series of 8 to 12 points almost, and not the longer ones).
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import type { EtsRecorded } from './ets-excel-compare';
import { detectSeasonality } from './functions/ets-season';

const load = (name: string): EtsRecorded[] =>
	JSON.parse(
		readFileSync(fileURLToPath(new URL(`./__fixtures__/${name}`, import.meta.url)), 'utf8'),
	) as EtsRecorded[];
const answer = (r: EtsRecorded): number => (r.results[0] as { value: number }).value;

const base = load('ets-detect-excel.json');
const variants = load('ets-detect-variants-excel.json');
const noise = base.filter((r) => r.name.startsWith('noise-'));

/** Rounds n / k to the nearest integer with halves going down, as Excel's noise periods do. */
const period = (n: number, k: number): number => Math.ceil(n / k - 0.5);

/** The Fourier frequency of the detrended series with the largest periodogram ordinate whose period is 2 or more with 2.5 cycles or more. */
function strongestBin(values: readonly number[]): number {
	const n = values.length;
	const mean = values.reduce((a, b) => a + b, 0) / n;
	const mx = (n - 1) / 2;
	let sxy = 0;
	let sxx = 0;
	values.forEach((v, i) => {
		sxy += (i - mx) * (v - mean);
		sxx += (i - mx) * (i - mx);
	});
	const z = values.map((v, i) => v - mean - (sxy / sxx) * (i - mx));
	let best = 0;
	let bestPower = -1;
	for (let k = 1; k <= n / 2; k++) {
		const m = period(n, k);
		if (m < 2 || n / m < 2.5) continue;
		let c = 0;
		let s = 0;
		z.forEach((v, i) => {
			c += v * Math.cos((2 * Math.PI * i * k) / n);
			s += v * Math.sin((2 * Math.PI * i * k) / n);
		});
		if (c * c + s * s > bestPower) {
			bestPower = c * c + s * s;
			best = k;
		}
	}
	return best;
}

describe('FORECAST.ETS.SEASONALITY against Excel on designed noise and waves', () => {
	it('records 352 series and 768 transformed copies', () => {
		expect(base.length).toBe(352);
		expect(variants.length).toBe(768);
	});

	it('is invariant to negation, reversal, rescaling and shifting in Excel', () => {
		const original = new Map(base.map((r) => [r.name, answer(r)]));
		const changed = variants.filter(
			(v) => answer(v) !== original.get(v.name.split('~')[0] as string),
		);
		expect(changed.map((v) => v.name)).toEqual([]);
	});

	it('is matched by this detector on at least 277 of the 352 series', () => {
		const same = base.filter((r) => detectSeasonality(r.values as number[]) === answer(r));
		expect(same.length).toBeGreaterThanOrEqual(277);
	});

	it('reports a period on 77 of the 256 noise series, at most 5 on those of up to 16 points', () => {
		const found = noise.filter((r) => answer(r) !== 0);
		expect(found.length).toBe(77);
		expect(Math.max(...found.filter((r) => r.values.length <= 16).map(answer))).toBeLessThanOrEqual(
			5,
		);
	});

	it('reports the period of the strongest periodogram bin on most noise series where it reports one', () => {
		const found = noise.filter((r) => answer(r) !== 0);
		const same = found.filter((r) => {
			const n = r.values.length;
			return period(n, strongestBin(r.values as number[])) === answer(r);
		});
		expect(same.length).toBeGreaterThanOrEqual(64);
	});
});
