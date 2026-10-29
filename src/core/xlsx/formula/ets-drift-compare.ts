// Labels every recorded non-seasonal probe of 8 or more points by what Excel forecast: the forecast of
// the (alpha, beta) pair it fits, or "drift" (the last value plus the least-squares slope, which
// FORECAST.ETS.STAT then reports as beta 0.001). Used by forecast-ets-drift.test.ts.
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { etsHoldout, etsRun } from './functions/ets-core';
import { initialStates, olsSlope } from './functions/ets-init';
import { CANDIDATES } from './functions/ets-search-map';
import { etsSearchSeries, etsValidationCost } from './functions/ets-validate';
import type { EtsRecorded } from './ets-excel-compare';

export interface DriftProbe {
	set: string;
	name: string;
	y: number[];
	/** The lowest hold-out cost pair of the lattice. */
	pair: readonly [number, number];
	/** Excel's STAT 2. */
	excelBeta: number;
	/** Excel's forecast and how many steps after the last point it is. */
	forecast: number;
	steps: number;
	/** Whether the forecast is the drift one (false: the pair's; null: neither within 1e-7). */
	drift: boolean | null;
	/** The pair filter's final trend and the least-squares slope. */
	trend: number;
	slope: number;
}

const SETS = ['search', 'morph', 'spike', 'scale', 'pairs', 'paths', 'short'];

function lowestPair(y: readonly number[]): readonly [number, number] {
	const hold = etsHoldout(y.length, 0);
	const searched = etsSearchSeries(y);
	const cost = (alpha: number, beta: number): number => {
		const p = { alpha, beta, gamma: 0.001 };
		return etsValidationCost(searched, 0, initialStates(y, 0, p, etsRun), p, hold);
	};
	let best: readonly [number, number] = [0.9, 0.001];
	let bestCost = cost(0.9, 0.001);
	for (const [alpha, beta] of CANDIDATES) {
		const c = cost(alpha, beta);
		if (c < bestCost * (1 - 1e-6) && c < bestCost - 1e-12) {
			bestCost = c;
			best = [alpha, beta];
		}
	}
	return best;
}

export function driftProbes(): DriftProbe[] {
	const out: DriftProbe[] = [];
	for (const set of SETS) {
		const path = fileURLToPath(new URL(`./__fixtures__/ets-${set}-excel.json`, import.meta.url));
		for (const r of JSON.parse(readFileSync(path, 'utf8')) as EtsRecorded[]) {
			const y = r.values as number[];
			if (y.length < 8 || y.some((v) => v === null)) continue;
			const beta = r.results.find((x) => /STAT\([^,]+,[^,]+,2,/.test(x.formula));
			const fc = r.results.find((x) => /^FORECAST\.ETS\(\d+,\{V\},\{L\},0\)$/.test(x.formula));
			if (beta?.type !== 'number' || fc?.type !== 'number') continue;
			const steps = Number(/ETS\((\d+)/.exec(fc.formula)?.[1]) - y.length;
			const pair = lowestPair(y);
			const state = (alpha: number, b: number) => {
				const p = { alpha, beta: b, gamma: 0.001 };
				const run = etsRun(y, 0, initialStates(y, 0, p, etsRun), p);
				return { forecast: run.level + steps * run.trend, trend: run.trend };
			};
			const own = state(pair[0], pair[1]);
			const drift = state(1, 0);
			const target = fc.value as number;
			const tol = 1e-7 * Math.max(1, Math.abs(target));
			const isPair = Math.abs(own.forecast - target) < tol;
			const isDrift = Math.abs(drift.forecast - target) < tol;
			out.push({
				set,
				name: r.name,
				y,
				pair,
				excelBeta: beta.value as number,
				forecast: target,
				steps,
				drift: isDrift && !isPair ? true : isPair && !isDrift ? false : null,
				trend: own.trend,
				slope: olsSlope(y),
			});
		}
	}
	return out;
}
