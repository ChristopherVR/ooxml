// Pure statistics on number lists, shared by the statistical functions, SUBTOTAL and AGGREGATE.
import { ERR, fail } from '../values.js';

export const sum = (values: readonly number[]): number => {
	let total = 0;
	let largest = 0;
	for (const v of values) {
		total += v;
		largest = Math.max(largest, Math.abs(v));
	}
	return total !== 0 && Math.abs(total) <= largest * 1e-15 ? 0 : total;
};

export function mean(values: readonly number[]): number {
	if (values.length === 0) fail(ERR.DIV0);
	return sum(values) / values.length;
}

/** Sum of squared deviations from the mean. */
export function devsq(values: readonly number[]): number {
	const m = mean(values);
	let total = 0;
	for (const v of values) total += (v - m) * (v - m);
	return total;
}

export function variance(values: readonly number[], sample: boolean): number {
	const n = values.length;
	if (n === 0 || (sample && n < 2)) fail(ERR.DIV0);
	return devsq(values) / (sample ? n - 1 : n);
}

export const stdev = (values: readonly number[], sample: boolean): number =>
	Math.sqrt(variance(values, sample));

export const sorted = (values: readonly number[]): number[] => [...values].sort((a, b) => a - b);

export function median(values: readonly number[]): number {
	if (values.length === 0) fail(ERR.NUM);
	const s = sorted(values);
	const mid = s.length >> 1;
	return s.length % 2 ? (s[mid] as number) : ((s[mid - 1] as number) + (s[mid] as number)) / 2;
}

/** The most frequent values in order of first appearance (MODE.MULT); `#N/A` without repeats. */
export function modes(values: readonly number[]): number[] {
	const counts = new Map<number, number>();
	let best = 1;
	for (const v of values) {
		const c = (counts.get(v) ?? 0) + 1;
		counts.set(v, c);
		best = Math.max(best, c);
	}
	if (best < 2) fail(ERR.NA);
	return [...counts].filter(([, c]) => c === best).map(([v]) => v);
}

/**
 * LARGE / SMALL: `k` must lie in 1..n before rounding; SMALL then truncates it and LARGE rounds
 * it up (LARGE(k) is SMALL(n + 1 - k)), as Excel does.
 */
export function kth(values: readonly number[], k: number, largest: boolean): number {
	const n = values.length;
	if (n === 0 || !(k >= 1) || k > n) fail(ERR.NUM);
	const index = Math.floor(largest ? n + 1 - k : k) - 1;
	return sorted(values)[index] as number;
}

/** The largest (or smallest) value without spreading the list into arguments; 0 when empty. */
export function extreme(values: readonly number[], max: boolean): number {
	if (values.length === 0) return 0;
	let best = values[0] as number;
	for (const v of values) if (max ? v > best : v < best) best = v;
	return best;
}

export function percentileInc(values: readonly number[], p: number): number {
	if (values.length === 0 || p < 0 || p > 1) fail(ERR.NUM);
	const s = sorted(values);
	const h = (s.length - 1) * p;
	const lo = Math.floor(h);
	const a = s[lo] as number;
	const b = s[Math.min(lo + 1, s.length - 1)] as number;
	return a + (h - lo) * (b - a);
}

export function percentileExc(values: readonly number[], p: number): number {
	const n = values.length;
	if (n === 0 || p <= 0 || p >= 1) fail(ERR.NUM);
	const h = (n + 1) * p;
	if (h < 1 || h > n) fail(ERR.NUM);
	const s = sorted(values);
	const lo = Math.floor(h);
	const a = s[lo - 1] as number;
	const b = s[Math.min(lo, n - 1)] as number;
	return a + (h - lo) * (b - a);
}

export function quartile(values: readonly number[], q: number, exclusive: boolean): number {
	const quart = Math.trunc(q);
	if (exclusive) {
		if (quart <= 0 || quart >= 4) fail(ERR.NUM);
		return percentileExc(values, quart / 4);
	}
	if (quart < 0 || quart > 4) fail(ERR.NUM);
	return percentileInc(values, quart / 4);
}

/** Excel's PERCENTRANK.INC: rank of `x` as a fraction, truncated to `significance` digits. */
export function percentRank(
	values: readonly number[],
	x: number,
	digits: number,
	exclusive: boolean,
): number {
	const s = sorted(values);
	const n = s.length;
	if (n === 0 || digits < 1) fail(ERR.NUM);
	const first = s[0] as number;
	const last = s[n - 1] as number;
	if (x < first || x > last) fail(ERR.NA);
	let rank: number;
	const lower = s.filter((v) => v < x).length;
	const exact = s.includes(x);
	if (exact) {
		rank = exclusive ? (lower + 1) / (n + 1) : n === 1 ? 1 : lower / (n - 1);
	} else {
		const below = s[lower - 1] as number;
		const above = s[lower] as number;
		const frac = (x - below) / (above - below);
		rank = exclusive ? (lower + frac) / (n + 1) : (lower - 1 + frac) / (n - 1);
	}
	const factor = Math.pow(10, Math.trunc(digits));
	return Math.floor(rank * factor + 1e-9) / factor;
}

export function covariance(xs: readonly number[], ys: readonly number[], sample: boolean): number {
	if (xs.length !== ys.length) fail(ERR.NA);
	const n = xs.length;
	if (n === 0 || (sample && n < 2)) fail(ERR.DIV0);
	const mx = mean(xs);
	const my = mean(ys);
	let total = 0;
	for (let i = 0; i < n; i++) total += ((xs[i] as number) - mx) * ((ys[i] as number) - my);
	return total / (sample ? n - 1 : n);
}

export function correlation(xs: readonly number[], ys: readonly number[]): number {
	const sx = Math.sqrt(devsq(xs));
	const sy = Math.sqrt(devsq(ys));
	if (sx === 0 || sy === 0) fail(ERR.DIV0);
	return (covariance(xs, ys, false) * xs.length) / (sx * sy);
}

/** Least-squares slope and intercept of y on x. */
export function linearFit(
	xs: readonly number[],
	ys: readonly number[],
): { slope: number; intercept: number } {
	if (xs.length !== ys.length) fail(ERR.NA);
	if (xs.length === 0) fail(ERR.DIV0);
	const mx = mean(xs);
	const my = mean(ys);
	let sxy = 0;
	let sxx = 0;
	for (let i = 0; i < xs.length; i++) {
		const dx = (xs[i] as number) - mx;
		sxy += dx * ((ys[i] as number) - my);
		sxx += dx * dx;
	}
	if (sxx === 0) fail(ERR.DIV0);
	const slope = sxy / sxx;
	return { slope, intercept: my - slope * mx };
}

/** RANK.EQ (or RANK.AVG with `average`): 1-based rank of `x` within `values`. */
export function rank(
	values: readonly number[],
	x: number,
	ascending: boolean,
	average: boolean,
): number {
	let better = 0;
	let ties = 0;
	for (const v of values) {
		if (v === x) ties++;
		else if (ascending ? v < x : v > x) better++;
	}
	if (ties === 0) fail(ERR.NA);
	return average ? better + (ties + 1) / 2 : better + 1;
}

const LANCZOS = [
	676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
	12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
];

/** The gamma function (Lanczos approximation). */
export function gamma(z: number): number {
	if (z < 0.5) return Math.PI / (Math.sin(Math.PI * z) * gamma(1 - z));
	const x = z - 1;
	let a = 0.99999999999980993;
	const t = x + 7.5;
	for (let i = 0; i < LANCZOS.length; i++) a += (LANCZOS[i] as number) / (x + i + 1);
	return Math.sqrt(2 * Math.PI) * Math.pow(t, x + 0.5) * Math.exp(-t) * a;
}

/** log(gamma(z)) for z > 0. */
export function gammaLn(z: number): number {
	if (z <= 0) fail(ERR.NUM);
	if (z < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * z))) - gammaLn(1 - z);
	const x = z - 1;
	let a = 0.99999999999980993;
	const t = x + 7.5;
	for (let i = 0; i < LANCZOS.length; i++) a += (LANCZOS[i] as number) / (x + i + 1);
	return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
}
