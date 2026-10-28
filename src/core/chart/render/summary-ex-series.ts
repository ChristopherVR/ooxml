import type { ChartExSeries, ChartExDataLabelOptions } from '../ex/model';
import type { ChartSummaryLabelLevel, ChartSummarySeries } from './summary';
const level = (source: ChartExDataLabelOptions): ChartSummaryLabelLevel => ({
	...(source.visibility?.value === undefined ? {} : { showValue: source.visibility.value }),
	...(source.visibility?.categoryName === undefined
		? {}
		: { showCategory: source.visibility.categoryName }),
	...(source.visibility?.seriesName === undefined
		? {}
		: { showSeriesName: source.visibility.seriesName }),
	...(source.position ? { position: source.position } : {}),
	...(source.numberFormat?.formatCode ? { numberFormat: source.numberFormat.formatCode } : {}),
	...(source.separator === undefined ? {} : { separator: source.separator }),
});
/** Paint and label facts for newly admitted extended-chart painters. Source relationships stay host-owned. */
export function chartExSeriesPaint(
	source: ChartExSeries,
	series: ChartSummarySeries,
	defaultValues = false,
): void {
	const fill = source.spPr?.fill;
	if (fill) series.fill = fill;
	if (source.spPr?.line) series.line = source.spPr.line;
	const points: NonNullable<ChartSummarySeries['pointColors']> = {};
	const fills: NonNullable<ChartSummarySeries['pointFills']> = {};
	for (const point of source.dataPoints) {
		if (point.index === undefined) continue;
		const fill = point.spPr?.fill;
		if (fill?.kind === 'solid') points[point.index] = fill.color;
		else if (fill) fills[point.index] = fill;
	}
	if (Object.keys(points).length) series.pointColors = points;
	if (Object.keys(fills).length) series.pointFills = fills;
	const labels = source.dataLabels;
	if (!labels && !defaultValues) return;
	const overrides: Record<number, ChartSummaryLabelLevel> = {};
	for (const label of labels?.labels ?? [])
		if (label.index !== undefined) overrides[label.index] = level(label);
	for (const index of labels?.hidden ?? [])
		overrides[index] = { ...overrides[index], deleted: true };
	series.dataLabels = {
		series: { showValue: defaultValues, ...(labels ? level(labels) : {}) },
		points: overrides,
	};
}
