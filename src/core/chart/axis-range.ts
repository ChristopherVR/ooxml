// Axis ranges and ticks extracted from the PowerPoint renderer.
import { computeAxisValueRange, computeAxisStackedValueRange } from './axis-range-auto';
import type { ChartAxisOptions, ChartAxisRange, ChartAxisSeries } from './axis-range-types';
type ValueRange = ChartAxisRange;

/** Compute a logarithmic range snapped to powers of the requested base. */
export function computeLogValueRange(
	series: ReadonlyArray<ChartAxisSeries>,
	logBase: number,
): ValueRange {
	const allValues = series
		.flatMap((item) => item.values)
		.filter((value): value is number => value !== null && Number.isFinite(value) && value > 0);
	if (allValues.length === 0) {
		return { min: 1, max: logBase, span: 1, logScale: true, logBase };
	}
	const dataMin = Math.min(...allValues);
	const dataMax = Math.max(...allValues);
	const logMin = Math.floor(Math.log(dataMin) / Math.log(logBase));
	const logMax = Math.ceil(Math.log(dataMax) / Math.log(logBase));
	const min = logBase ** logMin;
	const max = logBase ** Math.max(logMax, logMin + 1);
	return {
		min,
		max,
		span: Math.max(Math.log(max) / Math.log(logBase) - Math.log(min) / Math.log(logBase), 1),
		logScale: true,
		logBase,
	};
}

/**
 * Compute a value range using one parsed value axis's scale and bounds.
 *
 * @param plotHeightPx The chart's plot-area height, when known; forwarded to
 *   `computeValueRange`'s automatic-scale gridline count (see
 *   `axisTargetIntervals`'s doc comment). Not used on the log-scale branch,
 *   which snaps to powers of the log base regardless of chart size.
 */
export function computeValueRangeForAxis(
	series: ReadonlyArray<ChartAxisSeries>,
	axis: ChartAxisOptions | undefined,
	plotHeightPx?: number,
): ValueRange {
	const requestedLogBase = axis?.logBase ?? 10;
	const logBase =
		axis?.logScale && Number.isFinite(requestedLogBase) && requestedLogBase > 1
			? requestedLogBase
			: axis?.logScale
				? 10
				: undefined;
	const automatic = logBase
		? computeLogValueRange(series, logBase)
		: computeAxisValueRange(series, plotHeightPx);
	return withAxisBounds(automatic, axis, logBase);
}

/**
 * Compute a stacked chart's value range on its value axis: the category
 * totals, scaled by an explicit `c:min` / `c:max` the same way an unstacked
 * range is, or percentStacked's 0 to 100 (see {@link percentStackedRange}).
 * Both follow `c:orientation`.
 */
export function computeStackedValueRangeForAxis(
	series: ReadonlyArray<ChartAxisSeries>,
	catCount: number,
	axis: ChartAxisOptions | undefined,
	isPercent: boolean,
	plotSizePx?: number,
): ValueRange {
	if (isPercent) {
		return {
			...percentStackedRange(series, catCount, plotSizePx),
			...(axis?.orientation === 'maxMin' ? { reverseOrder: true } : {}),
		};
	}
	return withAxisBounds(computeAxisStackedValueRange(series, catCount, plotSizePx), axis);
}

/**
 * A percentStacked range runs from 0 to 100. A negative value takes its share
 * of the category's absolute total below zero, so when any share is negative
 * the bottom is the automatic scale of the lowest negative stack, never below
 * -100, and every segment stays on the plot.
 */
function percentStackedRange(
	series: ReadonlyArray<ChartAxisSeries>,
	catCount: number,
	plotSizePx?: number,
): ValueRange {
	const shares = series.map((entry) => ({
		...entry,
		values: Array.from({ length: catCount }, (_, ci) => {
			const total = series.reduce((sum, s) => sum + Math.abs(s.values[ci] ?? 0), 0);
			return total > 0 ? ((entry.values[ci] ?? 0) / total) * 100 : 0;
		}),
	}));
	const automatic = computeAxisStackedValueRange(shares, catCount, plotSizePx);
	if (automatic.min >= 0) {
		return { min: 0, max: 100, span: 100 };
	}
	const min = Math.max(automatic.min, -100);
	const unit = automatic.majorUnit;
	return {
		min,
		max: 100,
		span: 100 - min,
		...(unit !== undefined && 100 % unit === 0 && min % unit === 0 ? { majorUnit: unit } : {}),
	};
}

/**
 * Apply a value axis's explicit `c:min` / `c:max` and `c:orientation` to an
 * automatic range, stacked or not.
 */
export function withAxisBounds(
	automatic: ValueRange,
	axis: ChartAxisOptions | undefined,
	logBase?: number,
): ValueRange {
	const validMin =
		typeof axis?.min === 'number' && Number.isFinite(axis.min) && (!logBase || axis.min > 0)
			? axis.min
			: undefined;
	const validMax =
		typeof axis?.max === 'number' && Number.isFinite(axis.max) && (!logBase || axis.max > 0)
			? axis.max
			: undefined;
	const min = validMin ?? automatic.min;
	let max = validMax ?? automatic.max;
	// An explicit bound breaks the automatic scale's guarantee that the span is a
	// whole number of major units, so the unit no longer describes these bounds.
	const majorUnit =
		validMin === undefined && validMax === undefined ? automatic.majorUnit : undefined;
	if (max <= min) {
		max = logBase ? min * logBase : min + 1;
	}
	const span = logBase
		? Math.log(max) / Math.log(logBase) - Math.log(min) / Math.log(logBase)
		: max - min;
	return {
		min,
		max,
		span: Math.max(span, Number.EPSILON),
		...(majorUnit !== undefined ? { majorUnit } : {}),
		...(logBase ? { logScale: true, logBase } : {}),
		...(axis?.orientation === 'maxMin' ? { reverseOrder: true } : {}),
	};
}

/** Map a data value to a logarithmic Y coordinate. */
export function valueToYLog(val: number, range: ValueRange, topY: number, bottomY: number): number {
	const base = range.logBase ?? 10;
	const logValue = Math.log(Math.max(val, range.min)) / Math.log(base);
	const logMin = Math.log(range.min) / Math.log(base);
	const ratio = (logValue - logMin) / range.span;
	return range.reverseOrder ? topY + ratio * (bottomY - topY) : bottomY - ratio * (bottomY - topY);
}

/** Find any parsed logarithmic value axis. */
export function findLogAxis<T extends ChartAxisOptions>(axes: T[] | undefined): T | undefined {
	return axes?.find((axis) => axis.axisType === 'valAx' && axis.logScale);
}

/**
 * Compute the primary value range for a chart.
 *
 * @param plotHeightPx See `computeValueRangeForAxis`'s parameter of the same name.
 */
export function computeValueRangeForChart(
	series: ReadonlyArray<ChartAxisSeries>,
	axes?: ChartAxisOptions[],
	plotHeightPx?: number,
): ValueRange {
	const primaryAxis =
		axes?.find((axis) => axis.axisType === 'valAx' && axis.axPos === 'l') ??
		axes?.find((axis) => axis.axisType === 'valAx' && axis.axPos !== 'r');
	return computeValueRangeForAxis(series, primaryAxis, plotHeightPx);
}
