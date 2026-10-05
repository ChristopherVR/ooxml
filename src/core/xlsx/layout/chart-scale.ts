/** An automatic value-axis scale. */
export interface AxisScale {
	min: number;
	max: number;
	majorUnit: number;
	/** Tick values from `min` to `max` inclusive. */
	ticks: number[];
}

export interface NiceScaleOptions {
	/** Most major intervals allowed (Excel aims for about 5-10 depending on size; default 10). */
	maxIntervals?: number;
	/** Do not pull the axis to zero (scatter X axes). */
	noZero?: boolean;
}

const clean = (n: number): number => Number(n.toPrecision(12));

/** Candidate major units: 1, 2, 5 times powers of ten, from below `raw`. */
function* units(raw: number): Generator<number> {
	let exp = Math.floor(Math.log10(raw)) - 1;
	for (;;) {
		const base = 10 ** exp;
		for (const m of [1, 2, 5]) yield clean(m * base);
		exp++;
	}
}

/**
 * Excel-like automatic axis bounds: the axis includes zero unless the data sits in the top sixth
 * of its range, 5% headroom is added on the open ends, and the major unit is the smallest of
 * 1/2/5 x 10^n that keeps the interval count within `maxIntervals`.
 */
export function niceScale(
	dataMin: number,
	dataMax: number,
	options: NiceScaleOptions = {},
): AxisScale {
	const maxIntervals = Math.max(1, options.maxIntervals ?? 10);
	let lo = Number.isFinite(dataMin) ? dataMin : 0;
	let hi = Number.isFinite(dataMax) ? dataMax : 0;
	if (lo > hi) [lo, hi] = [hi, lo];
	if (lo === hi) {
		if (lo === 0) hi = 1;
		else if (options.noZero) {
			const pad = Math.abs(lo) * 0.1 || 1;
			lo -= pad;
			hi += pad;
		} else if (lo > 0) lo = 0;
		else hi = 0;
	}
	let fixedLo = false;
	let fixedHi = false;
	if (!options.noZero) {
		if (lo >= 0 && hi - lo >= hi / 6) {
			lo = 0;
			fixedLo = true;
		} else if (hi <= 0 && hi - lo >= -lo / 6) {
			hi = 0;
			fixedHi = true;
		}
	}
	const range = hi - lo;
	const top = fixedHi ? hi : hi + range * 0.05;
	const bottom = fixedLo ? lo : lo - range * 0.05;
	for (const unit of units(range / maxIntervals)) {
		const min = fixedLo ? lo : Math.floor(clean(bottom / unit)) * unit;
		const max = fixedHi ? hi : Math.ceil(clean(top / unit)) * unit;
		const intervals = Math.round((max - min) / unit);
		if (intervals <= maxIntervals && intervals > 0) {
			const ticks: number[] = [];
			for (let i = 0; i <= intervals; i++) ticks.push(clean(min + i * unit));
			return { min: clean(min), max: clean(max), majorUnit: unit, ticks };
		}
	}
	return { min: lo, max: hi, majorUnit: range, ticks: [lo, hi] };
}

/** The 0-100% scale of percent-stacked charts. */
export const PERCENT_SCALE: AxisScale = {
	min: 0,
	max: 1,
	majorUnit: 0.1,
	ticks: [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1],
};

/** Formats an axis value compactly (no float noise; percent axes as `40%`). */
export function formatAxisValue(value: number, percent = false): string {
	if (percent) return `${clean(value * 100)}%`;
	const v = clean(value);
	if (Math.abs(v) >= 1e15 || (v !== 0 && Math.abs(v) < 1e-6))
		return v.toExponential(2).replace('e+', 'E+').replace('e-', 'E-');
	const [int = '0', frac] = String(Math.abs(v)).split('.');
	const grouped = int.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
	return `${v < 0 ? '-' : ''}${grouped}${frac ? `.${frac}` : ''}`;
}
