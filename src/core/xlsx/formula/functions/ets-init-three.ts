// Excel's seasonal FORECAST.ETS start for a series of 3 whole seasons plus up to `ceil(m / 2) + 1`
// more points (seasons of 4 and more), and for exactly 2 seasons through the same rule. It was
// recovered exactly (to 1e-9) from the impulse probes of ets-init3 to ets-init12 (seasons 4 to 12;
// ets-init-three.test.ts). The 5-season chain rule is in ets-init-chain.ts; this is its 3-season
// relative, which is not a single chain but a blend of two:
//  * A step `e(t) = y[t] - y[t - 1] - B(t - j0)` with `j0 = floor(m / 2)` and `B(j)` the slope
//    `(y[j + 2m] - y[j]) / 2m` of position `j` (j >= m reads the position one season later, and is
//    used only while `j + 2m < n`; otherwise `B(j)` is the slope of position `j mod m`).
//  * Position 0 is the reference at time m. Position k is reached backwards (`back`: k at time k, up
//    to the reference) and forwards (`fwd`: the reference to k at time m + k). Positions below
//    `j0 - 1` are the forward path alone, and the last position is the backward path alone while
//    r <= ceil(m / 2) (r the points beyond 3 seasons); the others are `lev + (back + fwd) / 2` (the
//    level is the mean of the two ends) plus:
//      - `-(3u / 4m) * sum of the slopes of every position`, with u the number of positions the
//        forward path wraps past the season whose slope is unavailable (`u = k - j0 + 1 - r`, at
//        least 0), and slopes `S` the least-squares slopes of all of a position's values (a position
//        with a 4th value uses it here only),
//      - for each of those wrapped positions `i` the term `(y[i + m] + y[i + 2m] - 2 y[i]) / 4m`.
//  * The trend is the least-squares slope of the series, extended to `3m + ceil(m / 2)` points
//    by `y[t - 3m] + 3m * (mean of S)`, after removing the seasonal states.
//  * Two whole seasons: the series is extended to 3 seasons by `y[t - 2m] + 2m * b`, with b the slope
//    of a regression on the time and seasonal dummies, and then handled as above.
// Not covered (still approximate): 3 seasons with more than `ceil(m / 2) + 1` extra points (Excel's
// rule there is a different one, close to the 4-season chain: every position changes), 2 seasons
// plus a partial one, and seasons of 2 and 3 (their reference is time 0).

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** Least-squares slope of the values against 0, 1, 2, ... */
function slopeOf(v: readonly number[]): number {
	const k = v.length;
	const mx = (k - 1) / 2;
	const my = mean(v);
	let sxy = 0;
	let sxx = 0;
	for (let i = 0; i < k; i++) {
		sxy += (i - mx) * ((v[i] as number) - my);
		sxx += (i - mx) * (i - mx);
	}
	return sxx === 0 ? 0 : sxy / sxx;
}

/** Per-time-step slope of each position's values (every season that is present). */
function positionSlopes(y: readonly number[], m: number): number[] {
	return Array.from({ length: m }, (_v, c) => {
		const vals: number[] = [];
		for (let t = c; t < y.length; t += m) vals.push(y[t] as number);
		return slopeOf(vals) / m;
	});
}

/** Whether `threeSeasonSeasonals` is Excel's rule for a series of `n` values with seasons of `m`. */
export function isThreeSeasonLength(n: number, m: number): boolean {
	return m >= 4 && n >= 3 * m && n <= 3 * m + Math.ceil(m / 2) + 1;
}

/** Seasonal states relative to position 0 (`s[k] - s[0]`), see `isThreeSeasonLength`. */
export function threeSeasonSeasonals(y: readonly number[], m: number): number[] {
	const n = y.length;
	const extra = n - 3 * m;
	const j0 = Math.floor(m / 2);
	const at = (t: number): number => y[t] as number;
	const slope = (j: number): number => (at(j + 2 * m) - at(j)) / (2 * m);
	const B = (j: number): number => (j + 2 * m < n ? slope(j) : slope(j % m));
	const e = (t: number): number => at(t) - at(t - 1) - B(t - j0);
	const slopeSum = positionSlopes(y, m).reduce((a, b) => a + b, 0);
	const out = [0];
	for (let k = 1; k < m; k++) {
		let back = 0;
		for (let t = k + 1; t <= m; t++) back -= e(t);
		let fwd = 0;
		for (let t = m + 1; t <= m + k; t++) fwd += e(t);
		if (k < j0 - 1) {
			out.push(fwd);
			continue;
		}
		if (k === m - 1 && extra <= Math.ceil(m / 2)) {
			out.push(back);
			continue;
		}
		const wrapped = Math.max(0, k - j0 + 1 - extra);
		let corner = 0;
		for (let i = extra; i <= k - j0; i++)
			corner += (at(i + m) + at(i + 2 * m) - 2 * at(i)) / (4 * m);
		const level = (at(k) + at(m + k)) / 2 - at(m);
		const bb = back - (at(k) - at(m));
		const ff = fwd - (at(m + k) - at(m));
		out.push(level + (bb + ff) / 2 - ((3 * wrapped) / (4 * m)) * slopeSum + corner);
	}
	return out;
}

/** The series padded to `3m + ceil(m / 2)` points by `y[t - 3m] + 3m * (mean of S)`. */
export function padThreeSeasons(y: readonly number[], m: number): number[] {
	const n = y.length;
	const total = Math.max(3 * m + Math.ceil(m / 2), n);
	const slopeMean = mean(positionSlopes(y, m));
	const ext = [...y];
	for (let t = n; t < total; t++) ext.push((y[t - 3 * m] as number) + 3 * m * slopeMean);
	return ext;
}

/** The starting trend: slope of the series extended past 3 seasons, with the states removed. */
export function threeSeasonTrend(y: readonly number[], m: number, rel: readonly number[]): number {
	const ext = padThreeSeasons(y, m);
	const total = ext.length;
	const mx = (total - 1) / 2;
	let sxx = 0;
	let sxy = 0;
	for (let t = 0; t < total; t++) {
		sxx += (t - mx) * (t - mx);
		sxy += (t - mx) * ((ext[t] as number) - (rel[t % m] as number));
	}
	return sxy / sxx;
}

/** Whether a series of exactly 2 seasons is started through the 3-season rule. */
export function isTwoSeasonLength(n: number, m: number): boolean {
	return m >= 4 && n === 2 * m;
}

/** Two seasons extended to three by `y[t - 2m] + 2m * b` (b: slope with seasonal dummies). */
export function extendTwoSeasons(y: readonly number[], m: number): number[] {
	let num = 0;
	let den = 0;
	for (let c = 0; c < m; c++) {
		const t0 = c;
		const t1 = c + m;
		const mid = (t0 + t1) / 2;
		const avg = ((y[t0] as number) + (y[t1] as number)) / 2;
		num += (t0 - mid) * ((y[t0] as number) - avg) + (t1 - mid) * ((y[t1] as number) - avg);
		den += (t0 - mid) ** 2 + (t1 - mid) ** 2;
	}
	const b = num / den;
	return [...y, ...Array.from({ length: m }, (_v, c) => (y[c] as number) + 2 * m * b)];
}

/** The series a seasonal search measures its error over, or undefined where no padding applies. */
export function paddedSearchSeries(y: readonly number[], m: number): number[] | undefined {
	if (isTwoSeasonLength(y.length, m)) return padThreeSeasons(extendTwoSeasons(y, m), m);
	if (isThreeSeasonLength(y.length, m)) return padThreeSeasons(y, m);
	return undefined;
}
