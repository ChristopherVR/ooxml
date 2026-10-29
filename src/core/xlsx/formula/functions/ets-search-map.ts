// The (alpha, beta) lattice of Excel's FORECAST.ETS search without a season, and what Excel reports
// for the optimum on it (ets-search.ts has the evidence).
//
// The candidates are the 25 pairs Excel is seen to fit with: the error statistics of a series
// reproduce to the last digit with exactly one of them. The lowest hold-out cost over them is that
// pair on all 1392 earlier probes and on all 162 hard ones recorded for it (series chosen so that
// the cost is far lowest at a large-beta pair). Alpha 0.998 is only ever fitted with a beta of 0.099
// or more: with the lowest beta it is not a candidate (a penalty of 8% or more on it gives these
// results, of 5% six misses; of 10% the series of 8 to 10 points too). Pairs that are not Excel's
// own, tried earlier as stand-ins, are gone.
//
// What FORECAST.ETS.STAT reports is the pair's alpha always. For a large beta the beta it reports is
// the pair's, or 0.001 when Excel forecasts the last value plus the least-squares slope (alpha 1,
// beta 0, "drift") while the error statistics still use the pair. Which of the two happens is not
// fully known: pairs next to alpha - 0.001 of 0.1, 0.998 (0.749, 0.899) and the 1/6 pair always
// drift, the lowest beta never does, and for the others a series drifts
// when its hold-out mean lies within 1.5 standard deviations of the training mean and its slope is
// under 9.3 standard errors (83% of 643 probes, 72% by the commonest outcome of each pair).

export interface Outcome {
	/** The alpha FORECAST.ETS.STAT reports. */
	alpha: number;
	/** The beta it reports. */
	beta: number;
	/** The pair the error statistics are measured with. */
	fit: readonly [number, number];
	/** Excel forecasts from the last value plus the least-squares slope. */
	drift: boolean;
}

const LOW = 0.001;
/** Where the pair is not decided a series drifts when its hold-out mean is this close (in sd) to... */
const DRIFT_SHIFT = 1.5;
/** ...the training mean and its slope is not stronger than this many standard errors. */
const DRIFT_T = 9.3;

/** Candidates in the order they are tried (a later one wins only by a strictly lower cost). */
export const CANDIDATES: readonly (readonly [number, number])[] = [
	...[0.002, 0.1, 0.126, 0.25, 0.5, 0.75, 0.9].map((a) => [a, LOW] as const),
	[0.1, 0.099],
	[0.126, 0.125],
	[0.25, 0.249],
	[0.333333, 0.166667],
	[0.5, 0.25],
	[0.5, 0.499],
	[0.666667, 0.166667],
	[0.666667, 0.5],
	[0.75, 0.749],
	[0.833333, 0.75],
	[0.9, 0.899],
	[0.998, 0.099],
	[0.998, 0.249],
	[0.998, 0.499],
	[0.998, 0.749],
	[0.998, 0.899],
	[0.166667, 0.0833333],
	[0.833333, 0.0833333],
];

/** Pairs that drift in two thirds or more of the probes (10 of 10, 10 of 10, 3 of 3, 44 of 44, 62 of 69, 18 of 25, 9 of 11, 12 of 18). */
const ALWAYS_DRIFT = new Set([
	'0.1/0.099',
	'0.998/0.899',
	'0.998/0.749',
	'0.166667/0.0833333',
	'0.25/0.249',
	'0.9/0.899',
	'0.666667/0.5',
	'0.75/0.749',
]);
/** Pairs that are always reported as they are (0 of 4 drift). */
const NEVER_DRIFT = new Set(['0.333333/0.166667']);

/** What tells a drifting series from a trending one (see `reportedFor`). */
export interface TrendShape {
	/** |mean of the hold-out - mean of the training part| in standard deviations of the series. */
	shift: number;
	/** |t statistic| of the least-squares slope. */
	t: number;
}

/** The shape of `y` for a hold-out of `h` trailing points (a third of the series when h is 0). */
export function trendShape(y: readonly number[], h: number): TrendShape {
	const n = y.length;
	const hold = h > 0 ? h : Math.max(1, Math.round(0.3 * n));
	const mx = (n - 1) / 2;
	let my = 0;
	for (const v of y) my += v;
	my /= n;
	let sxy = 0;
	let sxx = 0;
	let syy = 0;
	for (let i = 0; i < n; i++) {
		const dy = (y[i] as number) - my;
		sxy += (i - mx) * dy;
		sxx += (i - mx) * (i - mx);
		syy += dy * dy;
	}
	if (syy === 0 || sxx === 0 || n < 3) return { shift: 0, t: 0 };
	const slope = sxy / sxx;
	const resid = syy - slope * sxy;
	const se = Math.sqrt(Math.max(resid, 0) / (n - 2) / sxx);
	let tail = 0;
	for (let i = n - hold; i < n; i++) tail += y[i] as number;
	let head = 0;
	for (let i = 0; i < n - hold; i++) head += y[i] as number;
	const shift = Math.abs(tail / hold - head / (n - hold)) / Math.sqrt(syy / n);
	return { shift, t: se === 0 ? Number.POSITIVE_INFINITY : Math.abs(slope / se) };
}

/** What Excel reports when the lowest cost of the lattice is at (alpha, beta). */
export function reportedFor(alpha: number, beta: number, shape: TrendShape): Outcome {
	const key = `${alpha}/${beta}`;
	let drift: boolean;
	if (beta === LOW) drift = false;
	else if (ALWAYS_DRIFT.has(key)) drift = true;
	else if (NEVER_DRIFT.has(key)) drift = false;
	else drift = shape.shift <= DRIFT_SHIFT && shape.t <= DRIFT_T;
	return { alpha, beta: drift ? LOW : beta, fit: [alpha, beta], drift };
}
