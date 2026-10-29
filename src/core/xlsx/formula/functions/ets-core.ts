// Additive exponential triple smoothing (ETS "AAA": additive error, trend and seasonality) behind
// the FORECAST.ETS family. Excel's optimiser is undocumented; this is a documented, deterministic
// equivalent calibrated on Excel's recorded answers (forecast-ets-excel.test.ts). Smoothing is in
// error-correction form:
//   forecast = l + b + s[t mod m]; e = y - forecast
//   l += b + alpha*e;  b += alpha*beta*e;  s[t mod m] += gamma*(1-alpha)*e
// The initial level, trend and seasonal states are not guessed: for any smoothing parameters the
// fitted values are linear in them, so they are solved by least squares (see `etsInit`), and the
// three parameters are then searched within Excel's bounds of 0.001 to 0.999.
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
	const gs = p.gamma * (1 - p.alpha);
	for (let t = 0; t < y.length; t++) {
		const k = m >= 2 ? t % m : 0;
		const f = level + trend + (m >= 2 ? (season[k] ?? 0) : 0);
		const e = (y[t] as number) - f;
		fitted.push(f);
		sse += e * e;
		level += trend + p.alpha * e;
		trend += p.alpha * p.beta * e;
		if (m >= 2) season[k] = (season[k] ?? 0) + gs * e;
	}
	return { fitted, level, trend, season, sse };
}

const LOW = 0.001;
const HIGH = 0.999;
const clamp = (v: number): number => Math.min(HIGH, Math.max(LOW, v));

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
	let best: number[] = [0.5, 0.1, 0.1];
	let bestCost = Number.POSITIVE_INFINITY;
	const grid = [LOW, 0.1, 0.3, 0.5, 0.7, 0.9, HIGH];
	for (const a of grid)
		for (const b of grid)
			for (const g of m >= 2 ? grid : [LOW]) {
				const c = cost([a, b, g]);
				if (c < bestCost * (1 - 1e-9)) {
					bestCost = c;
					best = [a, b, g];
				}
			}
	for (let step = 0.1; step > 1e-4; step /= 2) {
		let improved = true;
		while (improved) {
			improved = false;
			for (let i = 0; i < (m >= 2 ? 3 : 2); i++)
				for (const dir of [-1, 1]) {
					const trial = [...best];
					trial[i] = clamp((best[i] as number) + dir * step);
					const c = cost(trial);
					if (c < bestCost * (1 - 1e-9)) {
						bestCost = c;
						best = trial;
						improved = true;
					}
				}
		}
	}
	const [alpha, beta, gamma] = best as [number, number, number];
	const { run } = solve({ alpha, beta, gamma });
	return { alpha, beta, gamma: m >= 2 ? gamma : 0, m, n, ...run };
}

/** The point forecast `h` steps after the last point (h <= 0 reads the fitted value at n - 1 + h). */
export function etsPredict(fit: EtsFit, h: number): number {
	if (h <= 0) return fit.fitted[Math.max(0, fit.n - 1 + h)] ?? Number.NaN;
	const k = fit.m >= 2 ? (fit.n - 1 + h) % fit.m : 0;
	return fit.level + h * fit.trend + (fit.m >= 2 ? (fit.season[k] ?? 0) : 0);
}

/** Standard deviation of the h-step forecast error (Hyndman's ETS(A,A,A) variance formula). */
export function etsErrorSd(fit: EtsFit, h: number): number {
	const sigma2 = fit.sse / Math.max(1, fit.n);
	const hb = fit.alpha * fit.beta;
	const hg = fit.gamma * (1 - fit.alpha);
	let v = 1;
	for (let j = 1; j < Math.max(1, h); j++) {
		const c = fit.alpha + j * hb + (fit.m >= 2 && j % fit.m === 0 ? hg : 0);
		v += c * c;
	}
	return Math.sqrt(sigma2 * v);
}
