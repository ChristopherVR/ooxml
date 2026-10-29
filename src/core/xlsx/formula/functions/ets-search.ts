// How the smoothing parameters of FORECAST.ETS are chosen. Excel's search is undocumented. What was
// measured against Excel 16 (491 non-seasonal series in __fixtures__/ets-search-excel.json plus the
// 34 of ets-excel.json, see docs/xlsx-parity-status.md):
//   - the objective is the one-step squared error of the filter run over the whole series from the
//     least-squares trend (weights, windows, hold-out splits, absolute errors, multi-step errors and
//     other starting states all explain fewer of Excel's choices);
//   - Excel lands on a coarse lattice of values (below), not on the continuous minimum, and stops
//     early on many series, so its exact path is not reproduced: the lowest error on the lattice is
//     Excel's choice on about 4 series in 10 beyond 16 points and 7 in 10 on 8 to 16 points;
//   - when alpha lands on 0.998 with the lowest beta, the forecast is `last value + trend`, with the
//     least-squares slope untouched (alpha 1, beta 0), which is what Excel prints for those series.
// On 8 to 16 points a small relative preference per value, fitted to those series, raises the
// match (5-fold held-out: 111 of 161 against 83 without it); beyond 16 points it did not help.
import type { Smoothing } from './ets-core';

export interface Chosen extends Smoothing {
	/** Excel's forecast for this fit is the last value plus the least-squares slope (alpha 1, beta 0). */
	drift: boolean;
}

export const LOW = 0.001;
/** Series up to this length can move beta to alpha - 0.001 and have their own preferences. */
const SHORT_SERIES = 16;
/** Series of fewer points than this keep the lattice of the earlier fit (no Excel data beyond them). */
const MIN_FLAT_POINTS = 8;

/** Values of alpha Excel is seen to choose without a season (on 8 or more points). */
const FLAT_ALPHAS = [0.002, 0.1, 0.126, 0.25, 0.5, 2 / 3, 0.75, 0.9, 0.998];
/** Relative cost added to each alpha above on 8 to 16 points (fitted to Excel's choices). */
const SHORT_ALPHA_COST = [0.04, 0, 0, 0, -0.04, 0, -0.04, -0.04, 0.02];
/** The same for beta on 8 to 16 points: the lowest, then alpha - 0.001, then the fixed betas. */
const SHORT_BETA_COST = [0, 0.15, 0.05];
/** The betas Excel is seen to choose with alpha 0.998 on 8 to 16 points, beside the lowest. */
const TOP_BETAS = [0.099, 0.249, 0.499];

/** Values of alpha Excel is seen to choose on fewer than 8 points. */
const SHORT_ALPHAS = [0.002, 0.1, 0.126, 0.25, 0.5, 2 / 3, 0.75, 0.9];
/** The same with a season: the middle values sit 0.001 higher. */
const SEASON_ALPHAS = [0.002, 0.126, 0.251, 0.501, 2 / 3, 0.75, 0.9];
/** Values of gamma (besides 1 - alpha - 0.001) Excel is seen to choose. */
const SEASON_STEPS = [LOW, 0.125, 0.25, 0.5, 0.75, 0.9];

/** Whether a cost beats the best so far by a relative margin (ties keep the first). */
function better(c: number, best: number): boolean {
	return c < best * (1 - 1e-6) && c < best - 1e-12;
}

type Cost = (p: Smoothing) => number;

/** Seasonal series: the lowest squared error on the earlier lattice, beta up to alpha on 7 to 16 points. */
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

/** Fewer than 8 points: the lowest squared error on the earlier lattice. */
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

/** 8 to 16 points: beta may be alpha - 0.001 (or one of three values with alpha 0.998). */
function searchShort(cost: Cost): Chosen {
	let best: Smoothing = { alpha: 0.9, beta: LOW, gamma: LOW };
	let bestCost = cost(best) * (1 + (SHORT_ALPHA_COST[7] as number));
	FLAT_ALPHAS.forEach((a, i) => {
		const betas: [number, number][] =
			a > 0.99
				? [[LOW, 0], ...TOP_BETAS.map((b): [number, number] => [b, 2])]
				: [
						[LOW, 0],
						[a - LOW, 1],
					];
		for (const [b, k] of betas) {
			const c =
				cost({ alpha: a, beta: b, gamma: LOW }) *
				(1 + (SHORT_ALPHA_COST[i] as number) + (SHORT_BETA_COST[k] as number));
			if (better(c, bestCost)) {
				bestCost = c;
				best = { alpha: a, beta: b, gamma: LOW };
			}
		}
	});
	return { ...best, drift: best.alpha === 0.998 && best.beta === LOW };
}

/** Beyond 16 points, 0.9 is preferred to its neighbours unless they lower the error by this share. */
const LONG_NINE_PREFERENCE = 0.06;

/**
 * Beyond 16 points: the lowest squared error on the lattice with the lowest beta. Alpha 0.998 is
 * left out: Excel does choose it (14 of 290 series) but taking it as often as the error says
 * is wrong about as often as right, and costs smooth series such as exponential growth their fit.
 */
function searchLong(cost: Cost): Chosen {
	// An exact fit costs the same everywhere and Excel reports 0.9 for it, so that is the start.
	let best: Smoothing = { alpha: 0.9, beta: LOW, gamma: LOW };
	let bestCost = cost(best) * (1 - LONG_NINE_PREFERENCE);
	for (const a of FLAT_ALPHAS.filter((x) => x < 0.99)) {
		const c =
			cost({ alpha: a, beta: LOW, gamma: LOW }) * (a === 0.9 ? 1 - LONG_NINE_PREFERENCE : 1);
		if (better(c, bestCost)) {
			bestCost = c;
			best = { alpha: a, beta: LOW, gamma: LOW };
		}
	}
	return { ...best, drift: false };
}

/**
 * Chooses alpha, beta and gamma. `cost` is the squared one-step error of a parameter set;
 * `m` is the season length (0 for none) and `n` the number of points.
 */
export function chooseSmoothing(n: number, m: number, cost: Cost): Chosen {
	if (m >= 2) return searchSeasonal(n, cost);
	if (n < MIN_FLAT_POINTS) return searchTiny(n, cost);
	return n <= SHORT_SERIES ? searchShort(cost) : searchLong(cost);
}
