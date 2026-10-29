// Seasonality detection for the FORECAST.ETS family: the dominant period of the linearly detrended
// series, found as the strongest peak of its periodogram. Calibrated on Excel's answers for 15
// varied series (see forecast-ets-excel.test.ts): Excel picks the dominant spectral component,
// which is why a 4-point pattern whose second harmonic is stronger than its fundamental reports 2.

/** Longest tail of the series that is analysed; keeps the quadratic DFT bounded. */
const MAX_POINTS = 8192;
/** A period must repeat at least this often in the data. */
const MIN_CYCLES = 3;
/** Smallest share of the spectrum's power the peak must hold to count as seasonal. */
const MIN_SHARE = 0.3;

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
 * The season length of a series, or 0 when it has no dominant repeating pattern (the value
 * FORECAST.ETS.SEASONALITY returns then). Non-integer periods (n / k) are rounded.
 */
export function detectSeasonality(series: readonly number[]): number {
	const y = series.length > MAX_POINTS ? series.slice(series.length - MAX_POINTS) : series;
	const n = y.length;
	if (n < 2 * MIN_CYCLES) return 0;
	const r = detrend(y);
	const scale = Math.max(...y.map(Math.abs), 1);
	if (r.every((v) => Math.abs(v) <= 1e-12 * scale)) return 0;
	const cos = new Float64Array(n);
	const sin = new Float64Array(n);
	for (let j = 0; j < n; j++) {
		cos[j] = Math.cos((2 * Math.PI * j) / n);
		sin[j] = Math.sin((2 * Math.PI * j) / n);
	}
	let total = 0;
	let best = 0;
	let bestK = 0;
	for (let k = 1; k <= Math.floor(n / 2); k++) {
		let re = 0;
		let im = 0;
		for (let t = 0; t < n; t++) {
			const j = (k * t) % n;
			const v = r[t] as number;
			re += v * (cos[j] as number);
			im -= v * (sin[j] as number);
		}
		const power = re * re + im * im;
		total += power;
		if (k >= MIN_CYCLES && power > best) {
			best = power;
			bestK = k;
		}
	}
	if (bestK === 0 || !(best / total >= MIN_SHARE)) return 0;
	const period = Math.round(n / bestK);
	return period >= 2 ? period : 0;
}
