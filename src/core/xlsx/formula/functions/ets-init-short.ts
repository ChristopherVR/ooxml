// Excel's seasonal FORECAST.ETS start below 2 seasons. Excel ignores the season for the shortest series
// (no seasonal states and gamma reported as the machine epsilon, the model is the non-seasonal one)
// and uses it from `ceil(10 (m + 1) / 7)` points (`seasonStart`, probed for seasons of 2 to 16:
// ets-init14, 15, 17 and 18). From there up to 2 seasons minus 1 (seasons of 5 and more) the start is
// the 2-season rule of `ets-init-three.ts` on the series padded to 2 seasons by `y[t - m]` (the points
// past the data repeat the previous season, so their positions have no slope), with two changes
// recovered exactly from the impulse probes (ets-init15, 16 and 18, seasons 5 to 16):
//  * the positions that were padded, except the last one (`k = n - m .. m - 2`), get `slopeSum / 2`
//    added to their seasonal state, `slopeSum` being the sum of the positions' slopes of the padded
//    series;
//  * the starting trend gets `trendShift(m, n) * slopeSum / sxx` added, `sxx` the sum of squares about the
//    mean of the time axis of the series padded for the trend (`3m + ceil(m / 2)` points); the trend
//    does not see the state corrections above.
import {
	extendTwoSeasons,
	padThreeSeasons,
	slopeSumOf,
	threeSeasonSeasonals,
	threeSeasonTrend,
} from './ets-init-three';

/**
 * The first length at which Excel uses a season of `m` (shorter series are non-seasonal): the
 * smallest `n` with `m + 1 <= 0.7 n`. Matches the first seasonal length of 15 seasons (2 to 16).
 */
export function seasonStart(m: number): number {
	return Math.ceil((10 * (m + 1)) / 7);
}

/** Whether `shortSeasonStart` is Excel's rule: from `seasonStart` up to 2 seasons minus 1 point. */
export function isShortSeasonLength(n: number, m: number): boolean {
	return m >= 5 && n >= seasonStart(m) && n < 2 * m;
}

/** The series padded to 2 seasons by the previous season's values. */
export function padTwoSeasons(y: readonly number[], m: number): number[] {
	const p = [...y];
	for (let t = y.length; t < 2 * m; t++) p.push(p[t - m] as number);
	return p;
}

/**
 * The multiple of `slopeSum / sxx` added to the starting trend (see the header): with `j0 =
 * floor(m / 2)` and `u = 2m - 1 - n` the points short of 2 seasons minus one,
 * `(u^2 - (j0 - 2) u + 2 (j0 - 1)) / 4`. A quadratic in `n` fitted on 29 recorded lengths (seasons 5 to
 * 13, to 1e-9), which includes `(j0 - 1) / 2` at `u = 0`.
 */
function trendShift(m: number, n: number): number {
	const j0 = Math.floor(m / 2);
	const u = 2 * m - 1 - n;
	return (u * u - (j0 - 2) * u + 2 * (j0 - 1)) / 4;
}

/** Seasonal states relative to position 0 and the starting trend, see the header. */
export function shortSeasonStart(
	y: readonly number[],
	m: number,
): { rel: number[]; trend: number } {
	const n = y.length;
	const p = padTwoSeasons(y, m);
	const ss = slopeSumOf(p, m);
	const ext = extendTwoSeasons(p, m);
	const base = threeSeasonSeasonals(ext, m, ss);
	const total = 3 * m + Math.ceil(m / 2);
	const sxx = (total * (total * total - 1)) / 12;
	const trend = threeSeasonTrend(ext, m, base, ss / m) + (trendShift(m, n) * ss) / sxx;
	const rel = base.map((s, k) => (k >= n - m && k <= m - 2 ? s + ss / 2 : s));
	return { rel, trend };
}

/** The series a seasonal search measures its error over for `shortSeasonStart`. */
export function shortSearchSeries(y: readonly number[], m: number): number[] {
	const p = padTwoSeasons(y, m);
	return padThreeSeasons(extendTwoSeasons(p, m), m, slopeSumOf(p, m) / m);
}
