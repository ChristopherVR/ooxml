// How the smoothing parameters of FORECAST.ETS are chosen. Excel's search is undocumented. What was
// measured against Excel 16 (about 1100 non-seasonal probe series: 491 in __fixtures__/
// ets-search-excel.json, 484 morphing a random walk into noise, 136 spikes and steps, see
// docs/xlsx-parity-status.md):
//   - the objective is not the one-step error of the whole series. The series is split into a
//     training part and a hold-out (`etsHoldout` trailing points); the cost is the one-step squared
//     error of the training part plus the squared multi-step forecast errors of the hold-out
//     (`ets-validate.ts`). Starting from the least-squares trend, this single change moved the match
//     with Excel's alpha from 44% to 85% (a constant with a spike or step in the hold-out is the
//     clearest evidence: the choice stops depending on the data in the hold-out as the training part
//     becomes constant);
//   - the result is scale, offset and sign invariant, so there is no absolute tolerance;
//   - Excel lands on a coarse lattice of values (below), not on the continuous minimum, so the
//     lowest cost on the lattice is taken. Excel avoids the upper end: 0.998 is chosen only when it
//     beats the rest by about 5%;
//   - a beta above the lowest is chosen on about 4% of series: alpha - 0.001 and, with alpha 0.998,
//     0.099, 0.249 or 0.499, each a little less often than the plain cost says;
//   - when alpha lands on 0.998 with the lowest beta, the forecast is `last value + trend`, with the
//     least-squares slope untouched (alpha 1, beta 0), which is what Excel prints for those series.
// Variants of the cost that lose clearly (alpha agreement on the 491 probes, against 471 here):
// starting the training run from the slope of the training window only (251), the training error
// alone (225), a second split inside the training part (149 alone, 326 added), hold-out forecasts
// from the training end with the training-window slope (298) or the full-run trend (357). Beta
// moves off 0.001 are matched on 42 of 49 only by the full-data start used here. Continuous
// optimisers (golden section, Nelder-Mead) and lattice descents on this cost do not beat it.
// What is not reproduced: the choice of 1/6 on a family of 51 bimodal series, and the rare alphas of
// 1/3, 2/3 and 5/6.
import type { Smoothing } from './ets-core';

export interface Chosen extends Smoothing {
	/** Excel's forecast for this fit is the last value plus the least-squares slope (alpha 1, beta 0). */
	drift: boolean;
}

export const LOW = 0.001;
/** Series of fewer points than this keep the lattice of the earlier fit (no Excel data beyond them). */
const MIN_FLAT_POINTS = 8;

/** Values of alpha Excel is seen to choose without a season (on 8 or more points). */
const FLAT_ALPHAS = [0.002, 0.1, 0.126, 0.25, 0.5, 0.75, 0.9, 0.998];
/** Relative cost added to alpha 0.998 (Excel prefers 0.9 unless 0.998 is clearly better). */
const TOP_COST = 0.05;
/** Relative cost added to a beta of alpha - 0.001 (chosen with the alphas below). */
const TIED_BETA_COST = 0.15;
/** The alphas that may take beta = alpha - 0.001. */
const TIED_BETA_ALPHAS = [0.126, 0.25, 0.5, 0.75];
/** The betas Excel is seen to choose with alpha 0.998, beside the lowest. */
const TOP_BETAS = [0.099, 0.249, 0.499];
/** Other (alpha, beta) pairs Excel is seen to choose, with a small extra cost. */
const OTHER_PAIRS: [number, number][] = [
	[0.5, 0.25],
	[0.666667, 0.166667],
	[0.333333, 0.166667],
	[0.666667, 0.5],
];
const OTHER_PAIR_COST = 0.03;

/** Values of alpha Excel is seen to choose on fewer than 8 points. */
const SHORT_ALPHAS = [0.002, 0.1, 0.126, 0.25, 0.5, 2 / 3, 0.75, 0.9];
/** The same with a season: the middle values sit 0.001 higher. */
const SEASON_ALPHAS = [0.002, 0.126, 0.251, 0.501, 2 / 3, 0.75, 0.9];
/** Values of gamma (besides 1 - alpha - 0.001) Excel is seen to choose. */
const SEASON_STEPS = [LOW, 0.125, 0.25, 0.5, 0.75, 0.9];
/** Series up to this length can move beta to alpha - 0.001 with a season or without one below 8 points. */
const SHORT_SERIES = 16;

/** Whether a cost beats the best so far by a relative margin (ties keep the first). */
function better(c: number, best: number): boolean {
	return c < best * (1 - 1e-6) && c < best - 1e-12;
}

type Cost = (p: Smoothing) => number;

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

/** Fewer than 8 points: the lowest cost on the earlier lattice. */
function searchTiny(n: number, cost: Cost): Chosen {
	let best: Smoothing = { alpha: 0.9, beta: LOW, gamma: LOW };
	let bestCost = cost(best);
	for (const a of SHORT_ALPHAS)
		for (const b of n >= 7 ? [LOW, a - LOW] : [LOW]) {
			const c = cost({ alpha: a, beta: b, gamma: LOW });
			if (better(c, bestCost)) {
				bestCost = c;
				best = { alpha: a, beta: b, gamma: LOW };
			}
		}
	return { ...best, drift: false };
}

/** The (alpha, beta, extra relative cost) triples tried for 8 or more points without a season. */
function flatCandidates(): [number, number, number][] {
	const out: [number, number, number][] = [];
	for (const a of FLAT_ALPHAS) {
		out.push([a, LOW, a > 0.99 ? TOP_COST : 0]);
		if (TIED_BETA_ALPHAS.includes(a)) out.push([a, a - LOW, TIED_BETA_COST]);
	}
	for (const b of TOP_BETAS) out.push([0.998, b, TOP_COST - 0.02]);
	for (const [a, b] of OTHER_PAIRS) out.push([a, b, OTHER_PAIR_COST]);
	return out;
}
const FLAT_CANDIDATES = flatCandidates();

/** 8 or more points without a season: the cheapest candidate, ties keeping the earlier one. */
function searchFlat(cost: Cost): Chosen {
	// An exact fit costs the same everywhere and Excel reports 0.9 for it, so that is the start.
	let best: Smoothing = { alpha: 0.9, beta: LOW, gamma: LOW };
	let bestCost = cost(best);
	for (const [a, b, extra] of FLAT_CANDIDATES) {
		const c = cost({ alpha: a, beta: b, gamma: LOW }) * (1 + extra);
		if (better(c, bestCost)) {
			bestCost = c;
			best = { alpha: a, beta: b, gamma: LOW };
		}
	}
	return { ...best, drift: best.alpha === 0.998 && best.beta === LOW };
}

/**
 * Chooses alpha, beta and gamma. `cost` is the training-plus-hold-out squared error of a parameter
 * set (`ets-validate.ts`); `m` is the season length (0 for none) and `n` the number of points.
 */
export function chooseSmoothing(n: number, m: number, cost: Cost): Chosen {
	if (m >= 2) return searchSeasonal(n, cost);
	return n < MIN_FLAT_POINTS ? searchTiny(n, cost) : searchFlat(cost);
}
