// The view-model parts beyond a single plain group: the groups of a combination chart, the
// running-total steps of a waterfall and the bins of a histogram. The arithmetic (steps, bins) is
// the format-neutral code PowerPoint shares (`../waterfall`, `../histogram-binning`).
import { chartPaletteSeriesColor } from '../color-palettes';
import { aggregateByCategory, computeHistogramBins } from '../histogram-binning';
import { buildWaterfallSteps, type WaterfallStep } from '../waterfall';
import { autoSeriesColor } from './chart-colors';
import { niceScale, formatAxisValue } from './chart-scale';
import type { ChartGroupView, ChartViewModel, ValueAxisView } from './chart-view';
import { valueExtent } from './chart-value-axis';
import { stockView } from './chart-view-stock';
import { seriesView, type ChartRefValue, type SeriesContext } from './chart-view-series';
import type { ChartSummary, ChartSummarySeries } from './summary';

/** A waterfall's steps and the colours of its three bar roles. */
export interface WaterfallView {
	steps: WaterfallStep[];
	/** Dashed connectors join consecutive bars unless the chart hides them. */
	connectors: boolean;
	colors: { increase: string; decrease: string; total: string };
}

/** The scale of the groups plotted on one value axis: their extents, nicely rounded. */
function groupsAxis(groups: readonly ChartGroupView[], categories: number): ValueAxisView {
	let min = Infinity;
	let max = -Infinity;
	for (const group of groups) {
		const extent = valueExtent(group.type, group.grouping, group.series, categories);
		min = Math.min(min, extent.min);
		max = Math.max(max, extent.max);
	}
	return min === Infinity
		? { ...niceScale(0, 1), percent: false }
		: { ...niceScale(min, max), percent: false };
}

/** The groups a combination chart layers over its first one, and the secondary axis they need. */
export function chartExtraView<S extends ChartSummarySeries>(
	chart: ChartSummary<S>,
	model: ChartViewModel,
	ctx: SeriesContext<S>,
	total: number,
	categoriesOf: (cats: ChartRefValue[]) => void,
): void {
	let index = chart.series.length;
	const groups: ChartGroupView[] = [];
	for (const group of chart.extraGroups ?? []) {
		const series = group.series.map((s) => seriesView(ctx, s, index++, total, categoriesOf));
		const view: ChartGroupView = {
			type: group.chartType,
			grouping: group.grouping ?? (group.chartType === 'column' ? 'clustered' : 'standard'),
			series,
			secondary: group.secondary,
		};
		if (group.chartType === 'stock') view.stock = stockView(group.stock, ctx);
		if (group.barGapWidth !== undefined) view.barGapWidth = group.barGapWidth;
		if (group.barOverlap !== undefined) view.barOverlap = group.barOverlap;
		groups.push(view);
	}
	if (!groups.length) return;
	model.extraGroups = groups;
}

/** Scales the combination chart's axes once the categories are known. */
export function chartExtraAxes(model: ChartViewModel): void {
	const groups = model.extraGroups;
	if (!groups || !model.valueAxis) return;
	const count = model.categories.length;
	const secondary = groups.filter((g) => g.secondary);
	if (secondary.length) model.secondaryAxis = groupsAxis(secondary, count);
	const shared = groups.filter((g) => !g.secondary);
	if (!shared.length || model.valueAxis.percent) return;
	const own = valueExtent(model.type, model.grouping, model.series, count);
	let min = own.min;
	let max = own.max;
	for (const group of shared) {
		const extent = valueExtent(group.type, group.grouping, group.series, count);
		min = Math.min(min, extent.min);
		max = Math.max(max, extent.max);
	}
	model.valueAxis = { ...niceScale(min, max), percent: false };
}

/** Waterfall: one value series becomes running-total steps; the axis spans every step end. */
export function waterfallView<S extends ChartSummarySeries>(
	chart: ChartSummary<S>,
	model: ChartViewModel,
	ctx: SeriesContext<S>,
): void {
	const values = (model.series[0]?.values ?? []).map((v) => v ?? 0);
	const steps = buildWaterfallSteps(values, chart.waterfall);
	const role = (i: number): string =>
		(ctx.palette && chartPaletteSeriesColor(ctx.palette, i, 3, ctx.scheme)) ||
		autoSeriesColor(ctx.theme, i);
	model.waterfall = {
		steps,
		connectors: chart.waterfall?.connectorLines !== false,
		colors: { increase: role(0), decrease: role(1), total: role(2) },
	};
	const ends = steps.flatMap((step) => [step.startValue, step.endValue]);
	model.valueAxis = {
		...(ends.length ? niceScale(Math.min(...ends), Math.max(...ends)) : niceScale(0, 1)),
		percent: false,
	};
	const labels = model.categories;
	model.categories = steps.map((_, i) => labels[i] ?? String(i + 1));
}

/** Histogram: the raw observations (or categories to count) become bins of one count series. */
export function histogramView<S extends ChartSummarySeries>(
	chart: ChartSummary<S>,
	model: ChartViewModel,
): void {
	const first = model.series[0];
	if (!first) return;
	const raw = first.values.map((v) => v ?? Number.NaN);
	const options = chart.histogram ?? {};
	const bins = options.aggregateByCategory
		? aggregateByCategory(
				first.values.map((v) => v ?? 0),
				model.categories,
			)
		: computeHistogramBins(raw, options, formatAxisValue);
	model.categories = bins.map((bin) => bin.label);
	first.values = bins.map((bin) => bin.value);
	model.barGapWidth = 0;
	model.barOverlap = 0;
}
