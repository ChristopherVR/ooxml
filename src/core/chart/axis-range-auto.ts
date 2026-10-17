import { axisTargetIntervals, niceValueAxisBounds } from './axis-nice';
import type { ChartAxisRange, ChartAxisSeries } from './axis-range-types';

export function computeAxisValueRange(
	series: ReadonlyArray<ChartAxisSeries>,
	plotHeightPx?: number,
): ChartAxisRange {
	let dataMin = Number.POSITIVE_INFINITY,
		dataMax = Number.NEGATIVE_INFINITY;
	for (const item of series) {
		for (const value of item.values) {
			if (value !== null && Number.isFinite(value) && value < dataMin) {
				dataMin = value;
			}
			if (value !== null && Number.isFinite(value) && value > dataMax) {
				dataMax = value;
			}
		}
	}
	if (dataMin === Number.POSITIVE_INFINITY) {
		return { min: 0, max: 1, span: 1 };
	}
	const { min, max, majorUnit } = niceValueAxisBounds(
		dataMin,
		dataMax,
		plotHeightPx !== undefined ? axisTargetIntervals(plotHeightPx) : undefined,
	);
	return { min, max, span: Math.max(max - min, Number.EPSILON), majorUnit };
}

/**
 * Value range for a stacked bar: the per-category sums, then the same automatic
 * scale as any other value axis.
 *
 * @param plotHeightPx See `computeChartAxisRange`'s parameter of the same name.
 */
export function computeAxisStackedValueRange(
	series: ReadonlyArray<ChartAxisSeries>,
	catCount: number,
	plotHeightPx?: number,
): ChartAxisRange {
	let maxSum = 0,
		minSum = 0;
	for (let ci = 0; ci < catCount; ci++) {
		let pos = 0,
			neg = 0;
		for (const s of series) {
			const v = s.values[ci] ?? 0;
			if (v >= 0) {
				pos += v;
			} else {
				neg += v;
			}
		}
		maxSum = Math.max(maxSum, pos);
		minSum = Math.min(minSum, neg);
	}
	const { min, max, majorUnit } = niceValueAxisBounds(
		Math.min(minSum, 0),
		Math.max(maxSum, 0),
		plotHeightPx !== undefined ? axisTargetIntervals(plotHeightPx) : undefined,
	);
	return { min, max, span: Math.max(max - min, Number.EPSILON), majorUnit };
}
