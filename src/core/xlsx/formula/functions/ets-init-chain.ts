// Excel's starting state of the seasonal FORECAST.ETS model, recovered exactly from its impulse
// responses (ets-init-maps.ts; the laws are tested in ets-init-chain.test.ts). It holds when the
// series has exactly 4 whole seasons, or 5 and more (only the first 5 seasons are read and a partial
// season after them is ignored; recorded for seasons of 2 to 7), and takes no smoothing parameter:
//  * Position k of the season (times k, k + m, ...) has `P(t)`, the mean of its first H seasons
//    (t >= m reads the same positions one or more seasons later), and `B(j)`, the half-split slope
//    of position j: the mean of its last h seasons minus the mean of its first h, over the distance
//    between them; with C seasons, h = floor(C / 2) and H = C - h.
//  * The seasonal states follow each other along m consecutive times, from j0 - 1 to j0 + m - 2
//    with j0 = floor(m / 2): `s[t] = s[t - 1] + P(t) - P(t - 1) - B(t - j0)`. Position 0 sits at
//    time m (time 0 when m is 2 or 3) and the first fitted value is the first observation.
//  * The starting trend is the least-squares slope of the whole series after removing those
//    seasonal states, so it ignores any periodic pattern.
// 3 seasons plus up to ceil(m / 2) + 1 points and exactly 2 seasons are in `ets-init-three.ts` (also reached
// through `chainStart`); other lengths are not covered here (`ets-init.ts` falls back to least
// squares for them).
import type { EtsInit } from './ets-init';
import {
	extendTwoSeasons,
	isThreeSeasonLength,
	isTwoSeasonLength,
	threeSeasonSeasonals,
	threeSeasonTrend,
} from './ets-init-three';

const MAX_SEASONS = 5;

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** Seasonal states relative to position 0 (`s[k] - s[0]`), from the first `C * m` values. */
export function chainSeasonals(y: readonly number[], m: number, seasons: number): number[] {
	const h = Math.floor(seasons / 2);
	const first = seasons - h;
	const at = (t: number): number => y[t] as number;
	const meanFirst = (t: number): number =>
		mean(Array.from({ length: first }, (_v, c) => at(t + c * m)));
	const slope = (j: number): number => {
		const early = mean(Array.from({ length: h }, (_v, c) => at(j + c * m)));
		const late = mean(Array.from({ length: h }, (_v, c) => at(j + (seasons - h + c) * m)));
		return (late - early) / ((seasons - h) * m);
	};
	const j0 = Math.floor(m / 2);
	const state = new Map<number, number>([[j0 - 1, 0]]);
	for (let t = j0; t <= j0 + m - 2; t++)
		state.set(t, (state.get(t - 1) as number) + meanFirst(t) - meanFirst(t - 1) - slope(t - j0));
	const byPosition: number[] = new Array<number>(m).fill(0);
	for (const [t, s] of state) byPosition[t % m] = s;
	const origin = byPosition[0] as number;
	return byPosition.map((s) => s - origin);
}

/**
 * Seasonal states of 4 whole seasons plus a partial one: the mean over every window of 4 whole
 * seasons (they start at 0 to `n mod m`). Not Excel's rule, which is unknown for this length: it
 * agrees with it on the differences between most positions and is off on one position.
 */
function windowMeanSeasonals(y: readonly number[], m: number): number[] {
	const extra = y.length % m;
	const sums = new Array<number>(m).fill(0);
	for (let o = 0; o <= extra; o++) {
		const rel = chainSeasonals(y.slice(o, o + 4 * m), m, 4);
		const origin = rel[(m - o) % m] as number;
		for (let k = 0; k < m; k++)
			sums[k] = (sums[k] as number) + (rel[(k - o + m) % m] as number) - origin;
	}
	return sums.map((s) => s / (extra + 1));
}

/** Whether `chainStart` is Excel's exact rule for a series of `n` values (not 4 seasons plus a bit). */
export function isExactLength(n: number, m: number): boolean {
	const whole = Math.floor(n / m);
	return (
		whole >= MAX_SEASONS ||
		(whole === 4 && n % m === 0) ||
		isThreeSeasonLength(n, m) ||
		isTwoSeasonLength(n, m)
	);
}

/** The initial states from seasonal states relative to position 0 and the trend. */
function fromRelative(y0: number, rel: readonly number[], trend: number): EtsInit {
	// z[k] = level + season[k] + trend is the fitted value of position k at time 0.
	const combined = rel.map((s) => y0 + s - trend);
	const level = mean(combined);
	return { level, trend, season: combined.map((c) => c - level) };
}

/**
 * The start for a series of 4 or more seasons, for 3 seasons and up to `ceil(m / 2) + 1` more points and for
 * exactly 2 seasons (seasons of 4 and more), or undefined for the rest. It is Excel's own except for
 * 4 seasons plus a partial one (`isExactLength` tells which).
 */
export function chainStart(y: readonly number[], m: number): EtsInit | undefined {
	const n = y.length;
	const whole = Math.floor(n / m);
	if (isThreeSeasonLength(n, m) || isTwoSeasonLength(n, m)) {
		const series = isTwoSeasonLength(n, m) ? extendTwoSeasons(y, m) : y;
		const rel = threeSeasonSeasonals(series, m);
		return fromRelative(y[0] as number, rel, threeSeasonTrend(series, m, rel));
	}
	if (m < 2 || whole < 4) return undefined;
	const partial = whole < MAX_SEASONS && n % m !== 0;
	const rel = partial
		? windowMeanSeasonals(y, m)
		: chainSeasonals(y, m, Math.min(whole, MAX_SEASONS));
	const mx = (n - 1) / 2;
	let sxy = 0;
	let sxx = 0;
	const my = mean(y.map((v, t) => v - (rel[t % m] as number)));
	for (let t = 0; t < n; t++) {
		const d = (y[t] as number) - (rel[t % m] as number) - my;
		sxy += (t - mx) * d;
		sxx += (t - mx) * (t - mx);
	}
	return fromRelative(y[0] as number, rel, sxy / sxx);
}
