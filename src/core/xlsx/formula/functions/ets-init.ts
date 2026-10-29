// Initial states of the exponential smoothing model behind FORECAST.ETS. Excel's first fitted
// value is always the first observation, so the states before the first point satisfy
// `level + season[0] + trend = y[0]`. Without a season the trend starts at the least-squares slope
// of the whole series (checked against Excel for series of 7 or more points); with a season the
// trend and the other seasonal states are solved by least squares under that constraint, which
// lands close to (but not exactly on) Excel's own starting values.
import type { Smoothing } from './ets-core';

export interface EtsInit {
	level: number;
	trend: number;
	season: number[];
}

type Runner = (
	y: readonly number[],
	m: number,
	init: EtsInit,
	p: Smoothing,
) => { fitted: number[] };

/** Longest season whose states are solved exactly; longer ones start from first-cycle values. */
const MAX_EXACT_SEASON = 24;

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** The least-squares slope of the series against 0, 1, 2, ... */
export function olsSlope(y: readonly number[]): number {
	const n = y.length;
	const mx = (n - 1) / 2;
	const my = mean(y);
	let sxy = 0;
	let sxx = 0;
	for (let i = 0; i < n; i++) {
		sxy += (i - mx) * ((y[i] as number) - my);
		sxx += (i - mx) * (i - mx);
	}
	return sxx === 0 ? 0 : sxy / sxx;
}

/** Splits combined level-plus-season values into a level (their mean) and zero-sum seasonals. */
function split(trend: number, combined: readonly number[]): EtsInit {
	const level = mean(combined);
	return { level, trend, season: combined.map((c) => c - level) };
}

/** A start for long seasons: the first cycle, detrended by the slope between the first two. */
function heuristicStates(y: readonly number[], m: number): EtsInit {
	const trend = (mean(y.slice(m, 2 * m)) - mean(y.slice(0, m))) / m;
	return split(
		trend,
		Array.from({ length: m }, (_v, k) => (y[k] as number) - trend * (k + 1)),
	);
}

/** Solves `a x = b` by Gaussian elimination with partial pivoting (destroys its inputs). */
function solveSystem(a: number[][], b: number[]): number[] {
	const k = b.length;
	for (let i = 0; i < k; i++) {
		let p = i;
		for (let r = i + 1; r < k; r++)
			if (Math.abs((a[r] as number[])[i] as number) > Math.abs((a[p] as number[])[i] as number))
				p = r;
		[a[i], a[p]] = [a[p] as number[], a[i] as number[]];
		[b[i], b[p]] = [b[p] as number, b[i] as number];
		const row = a[i] as number[];
		const pivot = row[i] as number;
		if (Math.abs(pivot) < 1e-12) continue;
		for (let r = i + 1; r < k; r++) {
			const other = a[r] as number[];
			const f = (other[i] as number) / pivot;
			for (let c = i; c < k; c++) other[c] = (other[c] as number) - f * (row[c] as number);
			b[r] = (b[r] as number) - f * (b[i] as number);
		}
	}
	const x = new Array<number>(k).fill(0);
	for (let i = k - 1; i >= 0; i--) {
		let s = b[i] as number;
		for (let c = i + 1; c < k; c++) s -= ((a[i] as number[])[c] as number) * (x[c] as number);
		const pivot = (a[i] as number[])[i] as number;
		x[i] = Math.abs(pivot) < 1e-12 ? 0 : s / pivot;
	}
	return x;
}

/** The initial states for these smoothing parameters. */
export function initialStates(y: readonly number[], m: number, p: Smoothing, run: Runner): EtsInit {
	const y0 = y[0] as number;
	if (m < 2) {
		const trend = olsSlope(y);
		return { level: y0 - trend, trend, season: [] };
	}
	if (m > MAX_EXACT_SEASON) return heuristicStates(y, m);
	// Unknowns u = (trend, c1 .. c(m-1)) with c0 = y0 - trend, where c is level plus season. The
	// fitted values are affine in u, so the least-squares solution comes from the normal equations.
	const zeros = new Array<number>(y.length).fill(0);
	const unit = (trend: number, season: number[]): number[] =>
		run(zeros, m, { level: 0, trend, season }, p).fitted;
	const at = (k: number, v: number): number[] =>
		Array.from({ length: m }, (_x, i) => (i === k ? v : 0));
	const base = run(y, m, { level: 0, trend: 0, season: at(0, y0) }, p).fitted;
	const columns: number[][] = [unit(1, at(0, -1))];
	for (let k = 1; k < m; k++) columns.push(unit(0, at(k, 1)));
	const target = y.map((v, i) => v - (base[i] as number));
	const gram = columns.map((ci) =>
		columns.map((cj) => ci.reduce((s, v, i) => s + v * (cj[i] as number), 0)),
	);
	for (let i = 0; i < gram.length; i++)
		(gram[i] as number[])[i] = ((gram[i] as number[])[i] as number) + 1e-9;
	const rhs = columns.map((ci) => ci.reduce((s, v, i) => s + v * (target[i] as number), 0));
	const u = solveSystem(gram, rhs);
	const trend = u[0] as number;
	return split(trend, [y0 - trend, ...u.slice(1)]);
}
