// Resolves the data labels of a series into the text and placement the painters draw: the
// point > series > group cascade of `c:dLbls` over the series' categories and values.
import { composeDataLabelText, resolveDataLabelFlags } from '../data-label-text';
import { formatAxisValue } from './chart-scale';
import type { ChartSeriesView, ChartViewModel } from './chart-view';
import type { ChartSummaryDataLabels, ChartSummarySeries } from './summary';

/** One point's label: what it says, where it sits and the colour its number format asks for. */
export interface ChartLabelView {
	text: string;
	/** `c:dLblPos` of the point, else the series, else the group; absent for the chart's default. */
	position?: string;
	color?: string;
}

/** The labels of a series by point index; a hole means the point shows none. */
export type SeriesLabelViews = (ChartLabelView | undefined)[];

function labelsOf(
	source: ChartSummaryDataLabels,
	view: ChartSeriesView,
	categories: string[],
): SeriesLabelViews | undefined {
	const { chart, series, points } = source;
	if (series?.deleted === true || (chart?.deleted === true && !series)) return undefined;
	// A percentage is a share of the series' total, which only a pie or doughnut offers.
	const percentBase = view.values.reduce<number>(
		(sum, v) => sum + (v === null ? 0 : Math.abs(v)),
		0,
	);
	const out: SeriesLabelViews = [];
	let any = false;
	view.values.forEach((value, i) => {
		const point = points?.[i];
		if (value === null || point?.deleted === true) return;
		const position = point?.position ?? series?.position ?? chart?.position;
		const x = view.xValues?.[i];
		const category = x === undefined || x === null ? categories[i] : formatAxisValue(x);
		let built: { text: string; color?: string | undefined } | undefined;
		if (point?.text !== undefined) built = { text: point.text };
		else {
			const flags = resolveDataLabelFlags({
				...(point ? { point } : {}),
				...(series ? { series } : {}),
				...(chart ? { chart } : {}),
			});
			built = composeDataLabelText({
				flags,
				seriesName: view.name,
				category,
				value,
				numberFormat: point?.numberFormat ?? series?.numberFormat ?? chart?.numberFormat,
				percentBase,
				bubbleSize: view.sizes?.[i] ?? undefined,
				formatValue: (v) => formatAxisValue(v),
				formatBubbleSize: (size) => formatAxisValue(size),
			});
		}
		if (!built) return;
		any = true;
		out[i] = {
			text: built.text,
			...(position ? { position } : {}),
			...(built.color ? { color: built.color } : {}),
		};
	});
	return any ? out : undefined;
}

/** Gives every series of the model (and of its further groups) the labels its summary asks for. */
export function applyDataLabels(
	model: ChartViewModel,
	sources: readonly ChartSummarySeries[],
	extra: readonly { series: readonly ChartSummarySeries[] }[],
): void {
	const apply = (source: ChartSummarySeries | undefined, view: ChartSeriesView): void => {
		if (!source?.dataLabels) return;
		const labels = labelsOf(source.dataLabels, view, model.categories);
		if (labels) view.labels = labels;
	};
	model.series.forEach((view, i) => apply(sources[i], view));
	model.extraGroups?.forEach((group, g) =>
		group.series.forEach((view, i) => apply(extra[g]?.series[i], view)),
	);
}
