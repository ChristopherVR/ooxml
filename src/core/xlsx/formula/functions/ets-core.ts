// Additive exponential triple smoothing (ETS "AAA": additive error, trend and seasonality) behind
// the FORECAST.ETS family. The filter was identified against Excel 16 (see forecast-ets-excel.test.ts
// and docs/xlsx-parity-status.md): for every point,
//   forecast = level + trend + season[t mod m];  e = y - forecast
//   level += trend + alpha*e;  trend += beta*e;  season[t mod m] += gamma*e
// and the first forecast is exactly the first observation (`ets-init.ts`). Excel's search for alpha,
// beta and gamma is undocumented and stops at coarse values; this one minimises the squared one-step
// error within Excel's bounds of 0.001 to 0.999, which lands on or near Excel's choice.
import { ERR, fail } from '../values';
import { initialStates, type EtsInit } from './ets-init';

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

const LOW = 0.001;
/** Excel keeps beta at its lower bound on longer series; on 7 to 16 points it sometimes goes to alpha. */
const SHORT_SERIES = 16;
/** What Excel reports for gamma when the model has no season. */
const NO_SEASON_GAMMA = Number.EPSILON;
/** Values of alpha Excel is seen to choose. */
const ALPHAS = [0.002, 0.1, 0.126, 0.25, 0.5, 2 / 3, 0.75, 0.9];
/** The same with a season: the middle values sit 0.001 higher. */
const SEASON_ALPHAS = [0.002, 0.126, 0.251, 0.501, 2 / 3, 0.75, 0.9];
/** Values of gamma (besides 1 - alpha - 0.001) Excel is seen to choose. */
const SEASON_STEPS = [LOW, 0.125, 0.25, 0.5, 0.75, 0.9];

/** Fits alpha, beta and gamma by a coarse grid followed by shrinking pattern search. */
export function fitEts(y: readonly number[], seasonLength: number): EtsFit {
	const n = y.length;
	if (n < 2) fail(ERR.NUM);
	const m = seasonLength >= 2 && n >= 2 * seasonLength ? seasonLength : 0;
	const solve = (p: Smoothing) => {
		const init = initialStates(y, m, p, etsRun);
		return { init, run: etsRun(y, m, init, p) };
	};
	const cost = (v: readonly number[]): number =>
		solve({ alpha: v[0] as number, beta: v[1] as number, gamma: v[2] as number }).run.sse;
	// Excel's search settles on a coarse set of values (0.9 for an exact fit, whose cost is the
	// same everywhere), keeps beta <= alpha and gamma <= 1 - alpha, and never exceeds alpha 0.9
	// except on very short series. It does not always reach the best point (on a few series it stops
	// early), so the best point of the same kind of lattice is the closest reproducible match.
	let best: number[] = [0.9, LOW, LOW];
	let bestCost = cost(best);
	for (const a of m >= 2 ? SEASON_ALPHAS : ALPHAS)
		for (const b of n >= 7 && n <= SHORT_SERIES ? [LOW, a - LOW] : [LOW])
			for (const g of m >= 2 ? [...SEASON_STEPS, 1 - a - LOW] : [LOW]) {
				if (b < LOW || g > 1 - a - LOW + 1e-9) continue;
				const c = cost([a, b, g]);
				if (c < bestCost * (1 - 1e-6) && c < bestCost - 1e-12) {
					bestCost = c;
					best = [a, b, g];
				}
			}
	const [alpha, beta, gamma] = best as [number, number, number];
	const { run } = solve({ alpha, beta, gamma });
	return { alpha, beta, gamma: m >= 2 ? gamma : NO_SEASON_GAMMA, m, n, ...run };
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
