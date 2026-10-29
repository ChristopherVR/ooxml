// Recovers Excel's FORECAST.ETS starting states from the answers recorded in
// __fixtures__/ets-init-excel.json (series in ets-init-series.json, written by
// scripts/xlsx/ets-init-probes.ts). For every series the file holds Excel's smoothing parameters
// (STAT 1 to 3) and its forecasts at 2m + 1 horizons. A forecast is
// `c[(n - 1 + h) mod m] + h * trend` where `c = level + season` after the last point, so the
// final state follows from the forecasts; the filter is affine in its starting state, so running
// it on unit states and solving the linear system gives the starting state Excel used.
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { etsRun } from './functions/ets-core';
import { solveSystem } from './functions/ets-init';

export interface EtsInitCase {
	name: string;
	y: number[];
	/** Season length passed to Excel (0 for none). */
	m: number;
	alpha: number;
	beta: number;
	gamma: number;
	/** Excel's starting level plus season, per position (one value without a season). */
	c: number[];
	trend: number;
	/** First fitted value minus the first observation: 0 when the recovery is consistent. */
	firstError: number;
}

interface Recorded {
	name: string;
	values: number[];
	results: { value: number }[];
}

/** Least squares by the normal equations; the columns are the regressors. */
function leastSquares(columns: number[][], target: number[]): number[] {
	const gram = columns.map((a) =>
		columns.map((b) => a.reduce((s, v, i) => s + v * (b[i] as number), 0)),
	);
	const rhs = columns.map((a) => a.reduce((s, v, i) => s + v * (target[i] as number), 0));
	return solveSystem(gram, rhs);
}

export function recoverInit(name: string, y: number[], m: number, values: number[]): EtsInitCase {
	const n = y.length;
	const [alpha, beta, gamma] = values as [number, number, number];
	const forecasts = values.slice(3);
	const p = { alpha, beta, gamma };
	const k = m >= 2 ? m : 1;
	// Final combined state (c per phase, then trend) from the forecasts.
	const columns: number[][] = Array.from({ length: k }, (_v, j) =>
		forecasts.map((_f, i) => (m >= 2 ? ((n + i) % m === j ? 1 : 0) : 1)),
	);
	columns.push(forecasts.map((_f, i) => i + 1));
	const final = leastSquares(columns, forecasts);
	// The filter's response: final state as a function of the starting state (c per phase, trend).
	const zeros = new Array<number>(n).fill(0);
	const finalOf = (u: number[], series: number[]): number[] => {
		const start =
			m >= 2
				? { level: 0, trend: u[k] as number, season: u.slice(0, k) }
				: { level: u[0] as number, trend: u[1] as number, season: [] };
		const r = etsRun(series, m, start, p);
		return [...(m >= 2 ? r.season.map((s) => s + r.level) : [r.level]), r.trend];
	};
	const unit = (j: number): number[] => Array.from({ length: k + 1 }, (_v, i) => (i === j ? 1 : 0));
	const offset = finalOf(unit(-1), y);
	const response = Array.from({ length: k + 1 }, (_v, j) => finalOf(unit(j), zeros));
	const matrix = offset.map((_v, i) => response.map((col) => col[i] as number));
	const u = solveSystem(
		matrix,
		final.map((v, i) => v - (offset[i] as number)),
	);
	const c = u.slice(0, k);
	const trend = u[k] as number;
	return {
		name,
		y,
		m,
		alpha,
		beta,
		gamma,
		c,
		trend,
		firstError: (c[0] as number) + trend - (y[0] as number),
	};
}

/** Every series of the starting-state probes with Excel's recovered starting state. */
export function recoveredInits(): EtsInitCase[] {
	const path = fileURLToPath(new URL('./__fixtures__/ets-init-excel.json', import.meta.url));
	const data = JSON.parse(readFileSync(path, 'utf8')) as Recorded[];
	return data
		.filter((e) => /^(s\d+-|n\d+-)/.test(e.name))
		.map((e) => {
			const m = /^s(\d+)-/.exec(e.name);
			return recoverInit(
				e.name,
				e.values,
				m ? Number(m[1]) : 0,
				e.results.map((r) => r.value),
			);
		});
}

/** The detection probes: a series and the period Excel reports for it. */
export function recordedDetections(): { name: string; y: number[]; excel: number }[] {
	const path = fileURLToPath(new URL('./__fixtures__/ets-init-excel.json', import.meta.url));
	const data = JSON.parse(readFileSync(path, 'utf8')) as Recorded[];
	return data
		.filter((e) => e.name.startsWith('det-'))
		.map((e) => ({ name: e.name, y: e.values, excel: (e.results[0] as { value: number }).value }));
}
