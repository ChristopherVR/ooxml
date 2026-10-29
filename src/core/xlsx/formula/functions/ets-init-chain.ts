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
// 4 seasons plus a partial one is `ets-init-four.ts`; 3 seasons plus up to ceil(m / 2) + 1 points and exactly 2
// seasons are `ets-init-three.ts`, and 3 seasons plus more points up to 4 seasons `ets-init-threeplus.ts` (all
// reached through `chainStart`); other lengths are not covered here (`ets-init.ts` falls back to least squares).
import type { EtsInit } from './ets-init';
import {
	extendTwoSeasons,
	isThreeSeasonLength,
	isTwoPlusLength,
	isTwoSeasonLength,
	slopeSumOf,
	threeSeasonSeasonals,
	threeSeasonTrend,
} from './ets-init-three';

import { fourSeasonPlusSeasonals } from './ets-init-four';
import {
	isLowThreeLength,
	isThreePlusLength,
	lowSeasonSeasonals,
	threePlusSeasonals,
} from './ets-init-threeplus';
import { isShortSeasonLength, shortSeasonStart } from './ets-init-short';
import { chainSeasonals } from './ets-init-states';

export { chainSeasonals, chainStates } from './ets-init-states';

const MAX_SEASONS = 5;

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** Whether `chainStart` is Excel's exact rule for a series of `n` values. */
export function isExactLength(n: number, m: number): boolean {
	const whole = Math.floor(n / m);
	return (
		whole >= MAX_SEASONS ||
		whole === 4 ||
		isThreeSeasonLength(n, m) ||
		isThreePlusLength(n, m) ||
		isLowThreeLength(n, m) ||
		isTwoSeasonLength(n, m) ||
		isTwoPlusLength(n, m) ||
		isShortSeasonLength(n, m)
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
 * 2 seasons plus a partial one and seasons of 2 or 3 below 4 seasons (`isExactLength` tells which).
 */
export function chainStart(y: readonly number[], m: number): EtsInit | undefined {
	const n = y.length;
	const whole = Math.floor(n / m);
	if (isShortSeasonLength(n, m)) {
		const { rel, trend } = shortSeasonStart(y, m);
		return fromRelative(y[0] as number, rel, trend);
	}
	if (isThreeSeasonLength(n, m) || isTwoSeasonLength(n, m) || isTwoPlusLength(n, m)) {
		const short = isTwoSeasonLength(n, m) || isTwoPlusLength(n, m);
		const series = short ? extendTwoSeasons(y, m) : y;
		const rel = threeSeasonSeasonals(series, m, short ? slopeSumOf(y, m) : undefined);
		return fromRelative(
			y[0] as number,
			rel,
			threeSeasonTrend(series, m, rel, short ? slopeSumOf(y, m) / m : undefined),
		);
	}
	if (isLowThreeLength(n, m)) {
		const short = n < 3 * m;
		const series = short ? extendTwoSeasons(y, m) : y;
		const rel = lowSeasonSeasonals(series, m, short ? slopeSumOf(y, m) : undefined);
		const pad = short ? slopeSumOf(y, m) / m : undefined;
		return fromRelative(y[0] as number, rel, threeSeasonTrend(series, m, rel, pad));
	}
	if (isThreePlusLength(n, m)) {
		const rel = threePlusSeasonals(y, m);
		return fromRelative(y[0] as number, rel, threeSeasonTrend(y, m, rel));
	}
	if (m < 2 || whole < 4) return undefined;
	const partial = whole < MAX_SEASONS && n % m !== 0;
	const rel = partial
		? fourSeasonPlusSeasonals(y, m)
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
