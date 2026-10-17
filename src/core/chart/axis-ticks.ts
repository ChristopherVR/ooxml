import type { ChartAxisOptions, ChartAxisRange as ValueRange } from './axis-range-types';
const LOG_EXPONENT_TOLERANCE = 1e-12;

/** Generate one tick for each power of the range's logarithmic base. */
export function generateLogTicks(range: ValueRange): number[] {
	if (!range.logScale || !range.logBase) {
		return [];
	}
	const base = range.logBase;
	const logMin = Math.log(range.min) / Math.log(base);
	const logMax = Math.log(range.max) / Math.log(base);
	const firstExponent = Math.ceil(logMin - LOG_EXPONENT_TOLERANCE);
	const lastExponent = Math.floor(logMax + LOG_EXPONENT_TOLERANCE);
	if (lastExponent < firstExponent) {
		return range.min === range.max ? [range.min] : [range.min, range.max];
	}
	return Array.from(
		{ length: lastExponent - firstExponent + 1 },
		(_unused, index) => base ** (firstExponent + index),
	);
}

function generateIntervalTicks(range: ValueRange, unit: number): number[] {
	if (!Number.isFinite(unit) || unit <= 0 || range.logScale) {
		return [];
	}
	const tolerance = Math.max(1, Math.abs(range.min), Math.abs(range.max)) * 1e-12;
	const first = Math.ceil((range.min - tolerance) / unit);
	const last = Math.floor((range.max + tolerance) / unit);
	const count = Math.min(Math.max(last - first + 1, 0), 2000);
	return Array.from({ length: count }, (_, index) => (first + index) * unit);
}

/** Generate major tick values, honoring an explicit `c:majorUnit`. */
export function generateAxisTicks(
	range: ValueRange,
	axis: ChartAxisOptions | undefined,
	defaultIntervals: number,
): number[] {
	if (range.logScale && range.logBase) {
		return generateLogTicks(range);
	}
	const explicit = generateIntervalTicks(range, axis?.majorUnit ?? 0);
	if (explicit.length > 0) {
		return explicit;
	}
	// The automatic scale already picked a round step and snapped the bounds to
	// whole multiples of it, so stepping by that unit reproduces PowerPoint's
	// labels exactly. Dividing the span evenly instead is what put ticks on
	// 10.4 / 20.8 / 31.2 where PowerPoint writes 20 / 40 / 60.
	const automatic = generateIntervalTicks(range, range.majorUnit ?? 0);
	if (automatic.length > 0) {
		return automatic;
	}
	return Array.from(
		{ length: defaultIntervals + 1 },
		(_, index) => range.min + (range.span / defaultIntervals) * index,
	);
}

/** Generate minor gridline values, excluding positions occupied by major ticks. */
export function generateMinorAxisTicks(
	range: ValueRange,
	axis: ChartAxisOptions | undefined,
): number[] {
	const minor = generateIntervalTicks(range, axis?.minorUnit ?? 0);
	if (minor.length === 0) {
		return [];
	}
	const majorUnit = axis?.majorUnit;
	if (!majorUnit || !Number.isFinite(majorUnit) || majorUnit <= 0) {
		return minor;
	}
	const tolerance = Math.max(1, Math.abs(range.min), Math.abs(range.max)) * 1e-10;
	return minor.filter(
		(value) => Math.abs(value / majorUnit - Math.round(value / majorUnit)) > tolerance,
	);
}
