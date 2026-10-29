// Initial states of the exponential smoothing model behind FORECAST.ETS. Excel's first fitted
// value is always the first observation, so the states before the first point satisfy
// `level + season[0] + trend = y[0]`. Without a season the trend starts at the least-squares slope
// of the whole series for 3 and for 7 or more points, and for 4 to 6 points at the slope of the line
// through the data extended to 7 points (`shortSeriesTrend`); this matches Excel to the last digit
// for 3 to 16 points (ets-init-excel.test.ts, ets-init-maps.test.ts).
// With a season Excel's starting states are a fixed linear function of the data that does not
// depend on the smoothing parameters, is exact for any trend plus season, and keeps any periodic
// pattern out of the trend. For 4 whole seasons and for 5 or more it is known exactly and lives in
// `ets-init-chain.ts` (it reproduces Excel's forecasts at its own parameters to rounding
// error on about 1200 probe series, seasons 2 to 12); for 4 seasons plus a partial one that file
// uses a mean over windows, close but not exact. The laws found for shorter series are tested in
// ets-init-maps.test.ts: below 3 seasons Excel runs its 3-season algorithm on the series extended
// by copies of the first points moved by `n` slopes of a trend-plus-dummies fit; the 3-season
// algorithm itself is not known (for 3 seasons of 3 the first state follows the 4-season rule on the
// series extended the same way, the others do not), so for 3 seasons or fewer the trend and the
// other seasonal states here are solved by least squares under the first-point constraint, which
// lands near (not on) Excel's own starting values (a plain seasonal-dummy regression does worse on
// the recorded forecasts). Seasons of 25 and more start from first-cycle values when short.
// Eliminated against the recovered maps: backcasting and forward or backward passes with fixed
// smoothing parameters, geometric-weight regressions, iterated padding by y0 + k * slope or
// last + slope, constrained least squares of the filter at fixed parameters, classical
// decomposition, seasonal-dummy least squares with or without the first-point constraint, weighted
// dummy regressions with diagonal weights, the seasonal start of the hold-out training window, and
// for 4 seasons plus a partial one the first 4 seasons alone or any single window of them.
import type { Smoothing } from './ets-core';
import { chainStart } from './ets-init-chain';

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

/**
 * The points Excel appends to a series of 4 to 6 values before fitting the line, as `[source,
 * lag]` pairs for the times 4 (if the series has 4 values), 5 and 6: the point is `y[source] +
 * lag * slope`, with the slope the least-squares slope of the original values. Recovered exactly
 * from the starting trends Excel's forecasts imply (ets-init-excel.test.ts, 90 series; the exact
 * weights are (-8, -3, 2, 9) / 28 for 4 values, (-26, -13, -5, 13, 31) / 140 for 5 and
 * (-90, -124, -53, 18, 89, 160) / 980 for 6). The 6-value row is a copy of the first point six
 * steps on; the others are the unique decomposition of that form, not a derived rule.
 */
const PADS: Record<number, readonly (readonly [number, number])[]> = {
	6: [[0, 6]],
	5: [
		[1, 4],
		[0, 6],
	],
	4: [
		[0, 4],
		[1, 4],
		[0, 6],
	],
};

/**
 * Starting trend of a series without a season: the least-squares slope, except that 4 to 6 values
 * are first extended to 7 (`PADS`) and the line is fitted to all 7. Matches Excel to the last digit
 * for 3 to 16 values.
 */
export function shortSeriesTrend(y: readonly number[]): number {
	const pads = PADS[y.length];
	if (!pads) return olsSlope(y);
	const slope = olsSlope(y);
	return olsSlope([...y, ...pads.map(([src, lag]) => (y[src] as number) + lag * slope)]);
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
export function solveSystem(a: number[][], b: number[]): number[] {
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
		const trend = shortSeriesTrend(y);
		return { level: y0 - trend, trend, season: [] };
	}
	const exact = chainStart(y, m);
	if (exact) return exact;
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
