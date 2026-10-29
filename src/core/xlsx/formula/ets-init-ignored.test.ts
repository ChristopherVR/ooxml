// Below `seasonStart(m)` points Excel ignores the season (gamma is reported as the machine epsilon, the
// seasonal map of its starting state is exactly zero and the trend is the non-seasonal one), and from
// there it uses it. Recorded in ets-init14, 15, 16 (series with the impulse maps) and ets-init17 (three
// series for every length from the season plus 1 to 2 seasons minus 1, seasons of 2 to 16).
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { etsPredict, fitEts } from './functions/ets-core';
import { seasonStart } from './functions/ets-init-short';
import { recoveredMaps } from './ets-init-maps';
import { shortSeriesTrend } from './functions/ets-init';

interface Recorded {
	name: string;
	values: number[];
	results: { value: number }[];
}

function load(file: string): Recorded[] {
	const path = fileURLToPath(new URL(`./__fixtures__/${file}`, import.meta.url));
	return JSON.parse(readFileSync(path, 'utf8')) as Recorded[];
}

const EPSILON = Number.EPSILON;

describe('the season below its first length', () => {
	it('is used from ceil(10 (m + 1) / 7) points, on all 360 recorded series', () => {
		let checked = 0;
		for (const e of load('ets-init17-excel.json')) {
			const [, m, n] = /^c(\d+)-(\d+)-\d$/.exec(e.name) as RegExpExecArray;
			const gamma = (e.results[2] as { value: number }).value;
			const used = Math.abs(gamma - EPSILON) > 1e-12;
			expect(used, e.name).toBe(Number(n) >= seasonStart(Number(m)));
			checked++;
		}
		expect(checked).toBe(360);
	});

	it('has a zero seasonal map and the non-seasonal trend map', () => {
		// The map is read off Excel's forecasts, which carry no seasonal information at gamma = epsilon, so
		// a few groups recovered at extreme parameters are off: most must match.
		const ignored = recoveredMaps().filter(
			(g) => g.missing === 0 && g.m >= 2 && g.n < seasonStart(g.m) && g.n >= 3,
		);
		let good = 0;
		for (const g of ignored) {
			const zero = g.z.every((row) => row.every((v, i) => Math.abs(v - (i === 0 ? 1 : 0)) < 1e-6));
			const trend = g.trend.every((v, i) => {
				const y = Array.from({ length: g.n }, (_v, j) => (i === j ? 1 : 0));
				return Math.abs(v - shortSeriesTrend(y)) < 1e-6;
			});
			if (zero && trend) good++;
		}
		expect(ignored.length).toBeGreaterThan(20);
		expect(good / ignored.length).toBeGreaterThan(0.85);
	});

	it('is fitted as the non-seasonal model (gamma at epsilon, forecasts as Excel)', () => {
		let total = 0;
		let agree = 0;
		for (const file of [
			'ets-init15-excel.json',
			'ets-init16-excel.json',
			'ets-init17-excel.json',
		]) {
			for (const e of load(file)) {
				const mm = /^[sc](\d+)-(\d+)-/.exec(e.name);
				if (!mm || e.values.length >= seasonStart(Number(mm[1]))) continue;
				const fit = fitEts(e.values, Number(mm[1]));
				expect(fit.m).toBe(0);
				expect(fit.gamma).toBe(EPSILON);
				const v = e.results.map((r) => r.value);
				const first = file.includes('17') ? 4 : 3;
				const ex = v.slice(first);
				total++;
				const ok = ex.every(
					(x, h) => Math.abs(etsPredict(fit, h + 1) - x) <= 1e-9 * Math.max(1, Math.abs(x)),
				);
				if (ok) agree++;
			}
		}
		// The rest are non-seasonal series where the search lands on another lattice pair (as without a season).
		expect(agree / total).toBeGreaterThan(0.95);
	});
});
