// How the smoothing parameters of FORECAST.ETS are chosen. Excel's search is undocumented. What was
// measured against Excel 16 (about 1500 non-seasonal probe series: 491 in __fixtures__/
// ets-search-excel.json, 484 morphing a random walk into noise, 136 spikes and steps, 392 of the
// first rescaled, negated and reversed; docs/xlsx-parity-status.md has the detail):
//   - the objective is not the one-step error of the whole series. The series is split into a
//     training part and a hold-out (`etsHoldout` trailing points); the cost is the one-step squared
//     error of the training part plus the squared multi-step forecast errors of the hold-out
//     (`ets-validate.ts`), from the least-squares trend of the whole series. A weight other than 1
//     on the hold-out loses clearly (1188 of 1353 alphas at 1, 1125 at 0.9, 1089 at 1.1);
//   - the result is scale, offset and sign invariant, so there is no absolute tolerance;
//   - Excel searches a fixed lattice of 25 (alpha, beta) pairs, not the continuous optimum
//     (`ets-search-map.ts`). The error statistics of every probe are reproduced to the last digit by
//     one pair of the lattice, and the lowest cost on it is that pair on all of them: 1392 earlier
//     probes, 162 hard ones recorded for large-beta pairs, the 8 to 10 point series (120 recorded,
//     no hold-out: the plain one-step error) and the 4 to 7 point ones (extended to 7, see
//     `etsSearchSeries`). Alpha 0.998 is only fitted with a beta of 0.099 or more;
//   - what STAT 1 and 2 report is a function of that pair: its alpha always, and its beta unless the
//     forecast is the last value plus the least-squares slope (alpha 1, beta 0, "drift"), when STAT 2
//     says 0.001 while the error statistics still use the pair. Alpha 1/6, 2/3 and 5/6 are pairs
//     of their own ((1/6, 1/12) and (5/6, 1/12) were found by solving STAT 4-7 for the pair), not
//     rounded versions of another alpha. Whether a large-beta pair drifts is the one thing left
//     open: a few pairs always do, and for the rest a series drifts when its hold-out mean is within
//     1.5 sd of its training mean and its slope is under 9.3 standard errors (83% of 643 probes);
//     no single statistic does better (R squared 74%, cost ratios, trend ratios, n, h, margins,
//     autocorrelation, Mann-Kendall 60 to 79%), and blending a drifting and a pair series in 20
//     steps (231 recorded) flips the outcome at points where those statistics move smoothly.
// Variants of the cost that lose clearly (alpha agreement on the 491 probes): starting the training
// run from the slope of the training window only (251), the training error alone (225), a second
// split inside the training part (149 alone, 326 added), hold-out forecasts from the training end
// with the training-window slope (298) or the full-run trend (357), a mean error in place of the sum
// (1158), a rolling set of hold-out sizes or a skipped warm-up (no better than the one hold-out).
// Continuous optimisers (golden section, Nelder-Mead) and lattice descents on this cost do not
// beat it. A ratio threshold on the cost of 0.9 against 0.998 does not separate Excel's choices
// (they overlap from 1.001 to 1.057), so the choice is not a margin on that pair of costs.
// What is not reproduced: which large-beta pairs drift (see above).
import type { Smoothing } from './ets-core';
import { isExactLength } from './ets-init-chain';
import { CANDIDATES, reportedFor, type TrendShape } from './ets-search-map';

export interface Chosen extends Smoothing {
	/** Excel's forecast for this fit is the last value plus the least-squares slope (alpha 1, beta 0). */
	drift: boolean;
	/** The pair the error statistics are measured with when it is not the reported one. */
	fit?: Smoothing;
}

export const LOW = 0.001;
/** Series of fewer points than this report one fixed pair (no search is possible). */
const MIN_SEARCH_POINTS = 4;
/** What Excel reports for the 7 series of 3 points recorded, whatever their values. */
const THREE_POINTS: Chosen = { alpha: 0.1, beta: 0.099, gamma: LOW, drift: false };

/** Values of alpha Excel is seen to choose with a season (the middle values sit 0.001 higher). */
const SEASON_ALPHAS = [0.002, 0.126, 0.251, 0.501, 2 / 3, 0.75, 0.9];
/** Values of gamma (besides 1 - alpha - 0.001) Excel is seen to choose. */
const SEASON_STEPS = [LOW, 0.125, 0.25, 0.5, 0.75, 0.9];
/** Series up to this length can move beta to alpha - 0.001 with a season. */
const SHORT_SERIES = 16;

/** Whether a cost beats the best so far by a relative margin (ties keep the first). */
function better(c: number, best: number): boolean {
	return c < best * (1 - 1e-6) && c < best - 1e-12;
}

type Cost = (p: Smoothing) => number;

/**
 * Seasonal triples (alpha, beta, gamma) Excel is seen to report where the starting states are its
 * own (`isExactLength`). The lowest hold-out cost over exactly these (on the series padded as in
 * `ets-init-three.ts` where the start pads it) is Excel's triple on 2307 of the 2319 recorded
 * series that the forecasts identify (on the first 700, adding 0.002 or 0.9 with the lowest beta
 * and gamma dropped it to 506, so Excel does not try them; the last two triples, (0.251, 0.25,
 * 0.25) and (0.501, 0.5, 0.001), were added for the 3-season series). Reported as they are, with
 * no drift forecast.
 */
const SEASONAL_TRIPLES: readonly (readonly [number, number, number])[] = [
	[0.126, LOW, LOW],
	[0.251, LOW, LOW],
	[0.501, LOW, LOW],
	[0.751, LOW, LOW],
	[0.002, LOW, 0.125],
	[0.002, LOW, 0.25],
	[0.002, LOW, 0.5],
	[0.002, LOW, 0.75],
	[0.251, LOW, 0.25],
	[0.75, LOW, 0.249],
	[0.9, LOW, 0.099],
	[0.251, 0.25, LOW],
	[0.1, LOW, 0.899],
	[0.126, 0.125, LOW],
	[0.998, 0.099, LOW],
	[0.25, LOW, 0.749],
	[0.5, LOW, 0.499],
	[0.5, 0.25, LOW],
	[0.251, 0.25, 0.25],
	[0.501, 0.5, LOW],
];

/** Seasonal series with Excel's own starting states: the lowest cost on `SEASONAL_TRIPLES`. */
function searchSeasonalLattice(cost: Cost): Chosen {
	let best: Smoothing = { alpha: 0.126, beta: LOW, gamma: LOW };
	let bestCost = cost(best);
	for (const [alpha, beta, gamma] of SEASONAL_TRIPLES) {
		const c = cost({ alpha, beta, gamma });
		if (better(c, bestCost)) {
			bestCost = c;
			best = { alpha, beta, gamma };
		}
	}
	return { ...best, drift: false };
}

/** Seasonal series: the lowest cost on the earlier lattice, beta up to alpha on 7 to 16 points. */
function searchSeasonal(n: number, cost: Cost): Chosen {
	let best: Smoothing = { alpha: 0.9, beta: LOW, gamma: LOW };
	let bestCost = cost(best);
	for (const a of SEASON_ALPHAS)
		for (const b of n >= 7 && n <= SHORT_SERIES ? [LOW, a - LOW] : [LOW])
			for (const g of [...SEASON_STEPS, 1 - a - LOW]) {
				if (b < LOW || g > 1 - a - LOW + 1e-9) continue;
				const c = cost({ alpha: a, beta: b, gamma: g });
				if (better(c, bestCost)) {
					bestCost = c;
					best = { alpha: a, beta: b, gamma: g };
				}
			}
	return { ...best, drift: false };
}

/**
 * 4 or more points without a season: the lowest cost on the lattice of (alpha, beta), ties keeping
 * the earlier pair, then what Excel reports for that optimum (`ets-search-map.ts`).
 */
function searchLattice(shape: TrendShape, cost: Cost): Chosen {
	// An exact fit costs the same everywhere and Excel reports 0.9 for it, so that is the start.
	let best: readonly [number, number] = [0.9, LOW];
	let bestCost = cost({ alpha: 0.9, beta: LOW, gamma: LOW });
	for (const [alpha, beta] of CANDIDATES) {
		const c = cost({ alpha, beta, gamma: LOW });
		if (better(c, bestCost)) {
			bestCost = c;
			best = [alpha, beta];
		}
	}
	const out = reportedFor(best[0], best[1], shape);
	const [fa, fb] = out.fit;
	const same = fa === out.alpha && fb === out.beta;
	return {
		alpha: out.alpha,
		beta: out.beta,
		gamma: LOW,
		drift: out.drift,
		...(same ? {} : { fit: { alpha: fa, beta: fb, gamma: LOW } }),
	};
}

/**
 * Chooses alpha, beta and gamma. `cost` is the training-plus-hold-out squared error of a parameter
 * set (`ets-validate.ts`); `m` is the season length (0 for none) and `n` the number of points.
 */
export function chooseSmoothing(
	n: number,
	m: number,
	cost: Cost,
	shape: TrendShape = { shift: 9, t: 99 },
): Chosen {
	if (m >= 2) return isExactLength(n, m) ? searchSeasonalLattice(cost) : searchSeasonal(n, cost);
	if (n < MIN_SEARCH_POINTS) return THREE_POINTS;
	return searchLattice(shape, cost);
}
