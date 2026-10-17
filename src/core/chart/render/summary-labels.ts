// The data labels of a summary series: the three cascade levels of `c:dLbls` (group, series,
// point) as the painter needs them, read from a parsed chart part.
import type { ChartDataLabelOptions, ChartDataLabels, ChartSpaceSeries } from '../model-series';
import type { ChartSummaryDataLabels, ChartSummaryLabelLevel } from './summary';

function level(
	options: ChartDataLabelOptions,
	custom?: { text?: string | undefined },
): ChartSummaryLabelLevel {
	const out: ChartSummaryLabelLevel = {};
	if (options.deleted !== undefined) out.deleted = options.deleted;
	if (options.showValue !== undefined) out.showValue = options.showValue;
	if (options.showCategoryName !== undefined) out.showCategory = options.showCategoryName;
	if (options.showSeriesName !== undefined) out.showSeriesName = options.showSeriesName;
	if (options.showPercent !== undefined) out.showPercent = options.showPercent;
	if (options.showBubbleSize !== undefined) out.showBubbleSize = options.showBubbleSize;
	if (options.separator !== undefined) out.separator = options.separator;
	if (options.position) out.position = options.position;
	if (options.numberFormat?.formatCode) out.numberFormat = options.numberFormat.formatCode;
	if (custom?.text !== undefined) out.text = custom.text;
	return out;
}

/** Whether a level switches any label content on (or writes a custom text). */
const shows = (l: ChartSummaryLabelLevel | undefined): boolean =>
	l !== undefined &&
	(l.showValue === true ||
		l.showCategory === true ||
		l.showSeriesName === true ||
		l.showPercent === true ||
		l.showBubbleSize === true ||
		l.text !== undefined);

/**
 * The labels of one series: its own `c:dLbls` over the group's. Absent when neither level is
 * written or none of them switches anything on (Office writes an all-zero group level on every
 * chart it authors), which means the series shows no labels.
 */
export function summaryDataLabels(
	series: ChartSpaceSeries,
	group: ChartDataLabels | undefined,
): ChartSummaryDataLabels | undefined {
	const own = series.dataLabels;
	if (!own && !group) return undefined;
	const out: ChartSummaryDataLabels = {};
	if (group) out.chart = level(group);
	if (own) {
		out.series = level(own);
		const points: Record<number, ChartSummaryLabelLevel> = {};
		for (const label of own.labels)
			if (label.index !== undefined) points[label.index] = level(label, label.tx);
		if (Object.keys(points).length) out.points = points;
	}
	if (!shows(out.chart) && !shows(out.series) && !Object.values(out.points ?? {}).some(shows))
		return undefined;
	return out;
}
