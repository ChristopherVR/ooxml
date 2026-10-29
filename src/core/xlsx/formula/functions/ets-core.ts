// Additive exponential triple smoothing (ETS "AAA": additive error, trend and seasonality) behind
// the FORECAST.ETS family. The filter was identified against Excel 16 (see forecast-ets-excel.test.ts
// and docs/xlsx-parity-status.md): for every point,
//   forecast = level + trend + season[t mod m];  e = y - forecast
//   level += trend + alpha*e;  trend += beta*e;  season[t mod m] += gamma*e
// and the first forecast is exactly the first observation (`ets-init.ts`). Excel's search for alpha,
// beta and gamma is undocumented and stops at coarse values; `ets-search.ts` picks the minimum of
// the training-plus-hold-out error (`ets-validate.ts`) on the lattice Excel is seen to land on.
import { ERR, fail } from '../values';
import { initialStates, type EtsInit } from './ets-init';
import { isExactLength } from './ets-init-chain';
import { paddedSearchSeries } from './ets-init-three';
import { chooseSmoothing } from './ets-search';
import { trendShape } from './ets-search-map';
import { etsSearchSeries, etsValidationCost } from './ets-validate';

export interface EtsFit {
	alpha: number;
	beta: number;
	gamma: number;
	/** Season length; 0 when the model has no seasonal component. */
	m: number;
	/** One-step-ahead fitted values for every point of the series. */
	fitted: number[];
	level: number;
	trend: number;
	/** Seasonal state after the last point, indexed by `t mod m`. */
	season: number[];
	sse: number;
	n: number;
}

export interface Smoothing {
	alpha: number;
	beta: number;
	gamma: number;
}

/** Runs the filter. With `y` of zeros and a unit start state this is that state's response. */
export function etsRun(
	y: readonly number[],
	m: number,
	init: EtsInit,
	p: Smoothing,
): Omit<EtsFit, 'alpha' | 'beta' | 'gamma' | 'm' | 'n'> {
	let { level, trend } = init;
	const season = [...init.season];
	const fitted: number[] = [];
	let sse = 0;
	for (let t = 0; t < y.length; t++) {
		const k = m >= 2 ? t % m : 0;
		const f = level + trend + (m >= 2 ? (season[k] ?? 0) : 0);
		const e = (y[t] as number) - f;
		fitted.push(f);
		sse += e * e;
		level += trend + p.alpha * e;
		trend += p.beta * e;
		if (m >= 2) season[k] = (season[k] ?? 0) + p.gamma * e;
	}
	return { fitted, level, trend, season, sse };
}

/** What Excel reports for gamma when the model has no season. */
const NO_SEASON_GAMMA = Number.EPSILON;

/** Fits alpha, beta and gamma: the choice among Excel's lattice is in `ets-search.ts`. */
export function fitEts(y: readonly number[], seasonLength: number): EtsFit {
	const n = y.length;
	if (n < 2) fail(ERR.NUM);
	const m = seasonLength >= 2 && n >= 2 * seasonLength ? seasonLength : 0;
	const solve = (p: Smoothing) => {
		const init = initialStates(y, m, p, etsRun);
		return { init, run: etsRun(y, m, init, p) };
	};
	// Without a season Excel's cost is the training one-step error plus the hold-out forecast error
	// (`ets-validate.ts`). With a season it is the same where the starting states are Excel's own
	// (`ets-init-chain.ts` and `ets-init-three.ts`); elsewhere they are not, and that cost chooses
	// worse than the plain one-step error (season3-11: 8% against 1% forecast error). Where the
	// start pads the series (2 seasons, and 3 seasons plus up to `ceil(m / 2) + 1` points) the cost
	// runs over the padded series with the hold-out of its length: Excel's triple is then found on
	// 2307 of the 2319 recorded exact-start series, against 1790 on the unpadded one.
	const exact = m < 2 || isExactLength(n, m);
	const padded = m >= 2 && exact ? paddedSearchSeries(y, m) : undefined;
	const searched = padded ?? (m >= 2 ? y : etsSearchSeries(y));
	const holdout = m >= 2 && !exact ? 0 : etsHoldout(padded ? padded.length : n, m);
	const { drift, fit, ...chosen } = chooseSmoothing(
		n,
		m,
		(p) => etsValidationCost(searched, m, solve(p).init, p, holdout),
		trendShape(y, holdout),
	);
	const { run } = solve(fit ?? chosen);
	// Excel forecasts from the last value and the untouched slope where its pair drifts (ets-search-map.ts).
	const state = drift ? solve({ alpha: 1, beta: 0, gamma: chosen.gamma }).run : run;
	return {
		...chosen,
		gamma: m >= 2 ? chosen.gamma : NO_SEASON_GAMMA,
		m,
		n,
		...run,
		level: state.level,
		trend: state.trend,
	};
}

/**
 * How many trailing points Excel uses to measure the error statistics (none for fewer than 11
 * points): 40% of the series, at most the larger of 10, two seasons and 20% of it. Fitted to the
 * hold-outs of about 70 recorded series; the exact rule is not published.
 */
export function etsHoldout(n: number, m: number): number {
	if (n < 11) return 0;
	return Math.min(Math.floor(0.4 * n), Math.max(10, 2 * m, Math.round(0.2 * n)));
}

/** The point forecast `h` steps after the last point (h <= 0 reads the model line back from it). */
export function etsPredict(fit: EtsFit, h: number): number {
	const k = fit.m >= 2 ? (((fit.n - 1 + h) % fit.m) + fit.m) % fit.m : 0;
	return fit.level + h * fit.trend + (fit.m >= 2 ? (fit.season[k] ?? 0) : 0);
}

/** Standard deviation of the h-step forecast error (Hyndman's ETS(A,A,A) variance formula). */
export function etsErrorSd(fit: EtsFit, h: number): number {
	const sigma2 = fit.sse / Math.max(1, fit.n);
	const hb = fit.beta;
	const hg = fit.gamma;
	let v = 1;
	for (let j = 1; j < Math.max(1, h); j++) {
		const c = fit.alpha + j * hb + (fit.m >= 2 && j % fit.m === 0 ? hg : 0);
		v += c * c;
	}
	return Math.sqrt(sigma2 * v);
}
