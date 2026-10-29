// Seasonality detection for the FORECAST.ETS family: the first peak of the autocorrelation of the
// linearly detrended series that reaches 0.3, searching lags 2 to a third of the series. Fitted to
// Excel's answers on 122 series (see forecast-ets-excel.test.ts and ets-init-excel.json): 110
// agree, against 96 for the periodogram this replaced. It reproduces Excel on sinusoids of any
// period and length (including lengths that are not a multiple of the period, where a spectral
// peak rounds to the wrong period), on impulse trains with noise, and on noisy waves. What it does
// not reproduce is Excel's answers on pure noise: Excel reports a period on about a third of
// short noisy series (7 of 20 between 8 and 36 points, none beyond 48) and the statistic behind
// that is not the autocorrelation (it picks lags with autocorrelation below zero). 352 more designed
// series (forecast-ets-detect.test.ts) show it is symmetric in time and scale (it does not change
// when the series is negated, reversed or rescaled, so it is not a model fit) and that the period is
// n / k, rounded with halves down, for the strongest periodogram frequency on 66 of 77 noise
// series; no threshold on that peak's share decides when a period is reported, and a periodogram
// detector with a fitted threshold agrees on 384 of 463 series against 377 for this one.

/** Longest tail of the series that is analysed; keeps the quadratic autocorrelation bounded. */
const MAX_POINTS = 8192;
/** A period must repeat at least this often in the data. */
const MIN_CYCLES = 3;
/** Smallest autocorrelation a peak needs to count as seasonal. */
const MIN_ACF = 0.3;

const mean = (xs: readonly number[]): number => xs.reduce((a, b) => a + b, 0) / (xs.length || 1);

/** The series with its least-squares line removed. */
function detrend(y: readonly number[]): number[] {
	const n = y.length;
	const mx = (n - 1) / 2;
	const my = mean(y);
	let sxy = 0;
	let sxx = 0;
	for (let i = 0; i < n; i++) {
		sxy += (i - mx) * ((y[i] as number) - my);
		sxx += (i - mx) * (i - mx);
	}
	const slope = sxx === 0 ? 0 : sxy / sxx;
	return y.map((v, i) => v - my - slope * (i - mx));
}

/**
 * The season length of a series, or 0 when it has no repeating pattern (the value
 * FORECAST.ETS.SEASONALITY returns then).
 */
export function detectSeasonality(series: readonly number[]): number {
	const y = series.length > MAX_POINTS ? series.slice(series.length - MAX_POINTS) : series;
	const n = y.length;
	if (n < 2 * MIN_CYCLES) return 0;
	const r = detrend(y);
	const scale = Math.max(...y.map(Math.abs), 1);
	if (r.every((v) => Math.abs(v) <= 1e-12 * scale)) return 0;
	const maxLag = Math.floor(n / MIN_CYCLES);
	let energy = 0;
	for (const v of r) energy += v * v;
	const acf = (lag: number): number => {
		let s = 0;
		for (let t = 0; t + lag < n; t++) s += (r[t] as number) * (r[t + lag] as number);
		return s / energy;
	};
	let prev = acf(1);
	let cur = acf(2);
	for (let lag = 2; lag <= maxLag; lag++) {
		const next = acf(lag + 1);
		if (cur >= prev && cur >= next && cur >= MIN_ACF) return lag;
		prev = cur;
		cur = next;
	}
	return 0;
}
