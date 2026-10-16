// The length of a chart error bar in data units, for each ST_ErrValType: a fixed value, a
// percentage of the point, a multiple of the series' standard deviation, the standard error, or a
// custom amount per point. Standard deviation and error use the sample variance (`n - 1`), as
// Excel's STDEV.S and its chart error-bar options do.
import type { ChartErrorValueType } from './model-analysis';

/** What the amount of an error bar depends on (`ChartErrorBars`, resolved). */
export interface ChartErrorAmountOptions {
	valueType: ChartErrorValueType;
	/** `c:val`: the fixed value, the percentage or the standard-deviation multiple. */
	value?: number | undefined;
	/** Custom amounts by source point index. */
	customPlus?: readonly number[] | undefined;
	customMinus?: readonly number[] | undefined;
}

/**
 * The sample variance of `values` (`n - 1` denominator). A single value has none: the denominator
 * is clamped to 1 so the result is 0, not a division by zero.
 */
export function chartSampleVariance(values: readonly number[]): number {
	const count = values.length;
	if (count === 0) return 0;
	const mean = values.reduce((sum, value) => sum + value, 0) / count;
	return values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(count - 1, 1);
}

/**
 * The error amount of one point. `values` are the series' values along the bar's direction, in
 * display order; `displayIndex` is the point's place among them and `sourceIndex` its index in the
 * series (custom amounts are indexed by it); `side` is the half of the bar being drawn.
 */
export function chartErrorBarAmount(
	options: ChartErrorAmountOptions,
	values: readonly number[],
	displayIndex: number,
	sourceIndex: number,
	side: 'plus' | 'minus',
): number {
	switch (options.valueType) {
		case 'fixedVal':
			return options.value ?? 0;
		case 'percentage':
			return Math.abs(values[displayIndex] ?? 0) * ((options.value ?? 0) / 100);
		case 'stdDev':
			return Math.sqrt(chartSampleVariance(values)) * (options.value ?? 1);
		case 'stdErr':
			return Math.sqrt(chartSampleVariance(values) / Math.max(values.length, 1));
		case 'cust':
			return side === 'plus'
				? (options.customPlus?.[sourceIndex] ?? 0)
				: (options.customMinus?.[sourceIndex] ?? 0);
	}
}
