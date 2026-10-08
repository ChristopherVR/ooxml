import type { ChartObject, ChartSeries } from '../model';

/**
 * The series a chart edit targets. `index` arrives from callers at run time, so it must be an
 * own position of the series array before anything reads or writes through it: a key such as
 * `__proto__` or `constructor` would otherwise resolve to a prototype object, and the edit that
 * follows would mutate that prototype for every object. The lookup compares positions instead of
 * indexing with the caller's value, so no caller-chosen key ever selects the object written to.
 */
export function chartSeriesAt(chart: ChartObject, index: number): ChartSeries {
	const target = isSeriesPosition(chart, index)
		? chart.series.find((_, position) => position === index)
		: undefined;
	if (!target) throw new RangeError(`No chart series at index ${String(index)}`);
	return target;
}

/** Deep-clone the series list for a patch and return the clone's copy of the edited series. */
export function cloneChartSeriesForEdit(
	chart: ChartObject,
	index: number,
): { series: ChartSeries[]; target: ChartSeries } {
	chartSeriesAt(chart, index);
	const series = structuredClone(chart.series);
	return { series, target: series.find((_, position) => position === index)! };
}

function isSeriesPosition(chart: ChartObject, index: unknown): index is number {
	return (
		typeof index === 'number' &&
		Number.isSafeInteger(index) &&
		index >= 0 &&
		index < chart.series.length
	);
}
