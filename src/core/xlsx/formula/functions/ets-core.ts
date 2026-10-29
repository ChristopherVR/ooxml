// Additive exponential triple smoothing (ETS "AAA": additive error, trend and seasonality) behind
// the FORECAST.ETS family. Excel's own optimiser and seasonality detector are undocumented, so
// this is a documented, deterministic equivalent: results agree with Excel where the answer is
// determined (an exact linear trend, an exactly repeating pattern) and are otherwise close but
// not identical. Smoothing is in error-correction form:
//   forecast = l + b + s[t mod m]; e = y - forecast
//   l += b + alpha*e;  b += alpha*beta*e;  s[t mod m] += gamma*(1-alpha)*e
import { ERR, fail } from '../values';

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

interface Init {
	level: number;
	trend: number;
	season: number[];
}

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

function initialState(y: readonly number[], m: number): Init {
	if (m >= 2) {
		const first = y.slice(0, m);
		const second = y.slice(m, 2 * m);
		const trend = (mean(second) - mean(first)) / m;
		const detrended = first.map((v, j) => v - trend * (j + 1));
		const level = mean(detrended);
		return { level, trend, season: detrended.map((d) => d - level) };
	}
	const trend = (y[1] ?? 0) - (y[0] ?? 0);
	return { level: (y[0] ?? 0) - trend, trend, season: [] };
}

function run(
	y: readonly number[],
	m: number,
	init: Init,
	alpha: number,
	beta: number,
	gamma: number,
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
		level += trend + alpha * e;
		trend += alpha * beta * e;
		if (m >= 2) season[k] = (season[k] ?? 0) + gamma * (1 - alpha) * e;
	}
	return { fitted, level, trend, season, sse };
}

const clamp01 = (v: number): number => Math.min(1, Math.max(0, v));

/** Fits alpha, beta and gamma by a coarse grid followed by shrinking pattern search. */
export function fitEts(y: readonly number[], seasonLength: number): EtsFit {
	const n = y.length;
	if (n < 2) fail(ERR.NUM);
	const m = seasonLength >= 2 && n >= 2 * seasonLength ? seasonLength : 0;
	const init = initialState(y, m);
	const cost = (p: readonly number[]): number =>
		run(y, m, init, p[0] as number, p[1] as number, p[2] as number).sse;
	let best: number[] = [0.5, 0.1, 0.1];
	let bestCost = Number.POSITIVE_INFINITY;
	const grid = [0, 0.1, 0.3, 0.5, 0.7, 0.9, 1];
	for (const a of grid)
		for (const b of grid)
			for (const g of m >= 2 ? grid : [0]) {
				const c = cost([a, b, g]);
				if (c < bestCost - 1e-12) {
					bestCost = c;
					best = [a, b, g];
				}
			}
	for (let step = 0.1; step > 1e-5; step /= 2) {
		let improved = true;
		while (improved) {
			improved = false;
			for (let i = 0; i < (m >= 2 ? 3 : 2); i++)
				for (const dir of [-1, 1]) {
					const trial = [...best];
					trial[i] = clamp01((best[i] as number) + dir * step);
					const c = cost(trial);
					if (c < bestCost - 1e-12) {
						bestCost = c;
						best = trial;
						improved = true;
					}
				}
		}
	}
	const [alpha, beta, gamma] = best as [number, number, number];
	return { alpha, beta, gamma, m, n, ...run(y, m, init, alpha, beta, gamma) };
}

/** The point forecast `h` steps after the last point (h <= 0 reads the fitted value at n - 1 + h). */
export function etsPredict(fit: EtsFit, h: number): number {
	if (h <= 0) return fit.fitted[Math.max(0, fit.n - 1 + h)] ?? Number.NaN;
	const k = fit.m >= 2 ? (fit.n - 1 + h) % fit.m : 0;
	return fit.level + h * fit.trend + (fit.m >= 2 ? (fit.season[k] ?? 0) : 0);
}

/** Standard deviation of the h-step forecast error (Hyndman's ETS(A,A,A) variance formula). */
export function etsErrorSd(fit: EtsFit, h: number): number {
	const params = 2 + (fit.m >= 2 ? fit.m : 0);
	const sigma2 = fit.sse / Math.max(1, fit.n - params);
	const hb = fit.alpha * fit.beta;
	const hg = fit.gamma * (1 - fit.alpha);
	let v = 1;
	for (let j = 1; j < Math.max(1, h); j++) {
		const c = fit.alpha + j * hb + (fit.m >= 2 && j % fit.m === 0 ? hg : 0);
		v += c * c;
	}
	return Math.sqrt(sigma2 * v);
}

/**
 * Season length of a series: the smallest lag whose autocorrelation of the linearly detrended
 * series is near the strongest one and clearly positive, or 1 when there is no seasonality.
 */
export function detectSeasonality(y: readonly number[]): number {
	const n = y.length;
	if (n < 6) return 1;
	const xs = y.map((_v, i) => i);
	const mx = mean(xs);
	const my = mean(y);
	let sxy = 0;
	let sxx = 0;
	for (let i = 0; i < n; i++) {
		sxy += (i - mx) * ((y[i] as number) - my);
		sxx += (i - mx) * (i - mx);
	}
	const slope = sxx === 0 ? 0 : sxy / sxx;
	const r = y.map((v, i) => v - (my + slope * (i - mx)));
	const scale = Math.max(...y.map(Math.abs), 1);
	const energy = r.reduce((a, v) => a + v * v, 0);
	if (energy <= 1e-18 * scale * scale * n) return 1;
	const acf: number[] = [];
	const maxLag = Math.floor(n / 2);
	for (let lag = 2; lag <= maxLag; lag++) {
		let s = 0;
		for (let i = lag; i < n; i++) s += (r[i] as number) * (r[i - lag] as number);
		acf[lag] = (s / energy) * (n / (n - lag));
	}
	const top = Math.max(...acf.filter((v): v is number => v !== undefined));
	if (!(top > 0.5)) return 1;
	for (let lag = 2; lag <= maxLag; lag++) if ((acf[lag] as number) >= 0.9 * top) return lag;
	return 1;
}
