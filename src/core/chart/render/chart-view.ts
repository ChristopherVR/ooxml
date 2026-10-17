import { applyChartAxes, type ChartAxesView } from './chart-view-axes';
import type { ParentLabelLayout } from '../hierarchy';
import { bubbleView, hierarchyView } from './chart-view-families';
import { stockView, type StockView } from './chart-view-stock';
import type {
	BubbleSummaryOptions,
	ChartSummary,
	ChartSummaryGroupType,
	ChartSummarySeries,
	ChartSummaryType,
	ChartThemePalette,
} from './summary';
import { chartPaletteSeriesColor } from '../color-palettes';
import { autoSeriesColor } from './chart-colors';
import { niceScale, type AxisScale } from './chart-scale';
import { valueAxis } from './chart-value-axis';
export { categoryTotals } from './chart-value-axis';
import { chartAppearance, type ChartAppearance } from './chart-appearance';
import type { ChartGradientFill } from '../gradient-definition';
import type { DrawingSvgShadow } from '../../drawingml/drawing-shadow';
import { chartTitleText, type ChartTitleText } from './chart-title-text';
import type { ChartErrorBarsView, ChartTrendlineView } from './chart-view-analysis';
import type { ChartManualLayout } from '../manual-layout';
import { applyDataLabels, type SeriesLabelViews } from './chart-view-labels';
import {
	chartExtraAxes,
	chartExtraView,
	histogramView,
	waterfallView,
	type WaterfallView,
} from './chart-view-ex';
import {
	seriesContext,
	seriesView,
	toText,
	toNumber,
	type ChartRefValue,
	type ChartViewOptions,
} from './chart-view-series';
export type { ChartRefValue, ChartViewOptions, EvaluateRef } from './chart-view-series';
export type { WaterfallView } from './chart-view-ex';

export interface ChartSeriesView {
	shadow?: DrawingSvgShadow;
	shadowFilter?: string;
	gradient?: ChartGradientFill;
	pointGradients?: Record<number, ChartGradientFill>;
	name: string;
	values: (number | null)[];
	/** Scatter and bubble X values (numeric categories, or 1..n when the categories are text). */
	xValues?: (number | null)[];
	/** Bubble sizes, parallel to `values`. */
	sizes?: (number | null)[];
	color: string;
	/** Per-point colours (pie and doughnut vary colours by point). */
	pointColors?: string[];
	/** Data labels by point index. */
	labels?: SeriesLabelViews;
	/** Trendlines drawn over the series. */
	trendlines?: ChartTrendlineView[];
	/** Error bars drawn at each point. */
	errorBars?: ChartErrorBarsView[];
}

export interface ValueAxisView extends AxisScale {
	/** Ticks are fractions shown as percentages (percent-stacked charts). */
	percent: boolean;
}

/** A combination chart's further group: its series over the shared categories. */
export interface ChartGroupView {
	type: ChartSummaryGroupType;
	grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	barGapWidth?: number;
	barOverlap?: number;
	series: ChartSeriesView[];
	/** Stock groups: their resolved formatting. */
	stock?: StockView;
	/** Plotted against {@link ChartViewModel.secondaryAxis}. */
	secondary: boolean;
}

/** Neutral data for a chart painter: series resolved, colours chosen, axes scaled. */
export interface ChartViewModel {
	axes?: ChartAxesView;
	legendLayout?: ChartManualLayout;
	legendOverlay?: boolean;
	titleLayout?: ChartManualLayout;
	titleOverlay?: boolean;
	titleText?: ChartTitleText;
	barGapWidth?: number;
	barOverlap?: number;
	/** Native defaults and direct element formatting, resolved against the workbook theme. */
	appearance?: ChartAppearance;
	type: ChartSummaryType;
	grouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
	title?: string;
	showLegend: boolean;
	legendPosition: 'r' | 'l' | 't' | 'b' | 'tr';
	categories: string[];
	series: ChartSeriesView[];
	/** Value (Y) axis; for bar charts it runs horizontally. */
	valueAxis?: ValueAxisView;
	/** Combination chart: the groups layered over the first one. */
	extraGroups?: ChartGroupView[];
	/** Combination chart: the scale of the groups plotted on the secondary (right) value axis. */
	secondaryAxis?: ValueAxisView;
	/** Waterfall: the running-total steps and their bar roles. */
	waterfall?: WaterfallView;
	/** Stock: the resolved wick and bar formatting. */
	stock?: StockView;
	/** Bubble: the display options, defaults resolved. */
	bubble?: BubbleSummaryOptions;
	/** Treemap and sunburst: category levels (parents filled down, leaf first) and legend branches. */
	hierarchy?: {
		levels: string[][];
		parentLabelLayout: ParentLabelLayout;
		/** Whether cells and wedges carry their labels. */
		labels: boolean;
		branches: string[];
	};
	/** Scatter and bubble X axis. */
	xAxis?: AxisScale;
	/** Bars run horizontally (`bar`), not vertically (`column`). */
	horizontal: boolean;
	/** False for chart types this painter does not draw (surface). */
	supported: boolean;
}

const SUPPORTED: ReadonlySet<ChartSummaryType> = new Set([
	'bar',
	'column',
	'line',
	'area',
	'pie',
	'doughnut',
	'scatter',
	'radar',
	'waterfall',
	'histogram',
	'bubble',
	'stock',
	'treemap',
	'sunburst',
]);

/**
 * Resolves a chart's series, from live values when the host gives `evaluateRef` (the values of a
 * reference such as `Sheet1!$B$2:$B$9`, flattened row-major), falling back to the values cached
 * in the chart part, and chooses their colours from the theme.
 */
export function chartSummaryView<S extends ChartSummarySeries>(
	chart: ChartSummary<S>,
	theme: ChartThemePalette,
	options: ChartViewOptions<S> = {},
): ChartViewModel {
	const type = chart.chartType;
	const radial = type === 'pie' || type === 'doughnut';
	// Pie, treemap and sunburst colour by point (branch) and have no value axis.
	const byPoint = radial || type === 'treemap' || type === 'sunburst';
	const ctx = seriesContext(chart, theme, options);
	const { palette, scheme, drawingColor } = ctx;
	const grouping =
		chart.grouping ?? (type === 'bar' || type === 'column' ? 'clustered' : 'standard');
	const rawCategories: ChartRefValue[][] = [];
	const collect = (cats: ChartRefValue[]): void => void rawCategories.push(cats);
	const extra = chart.extraGroups ?? [];
	const total = chart.series.length + extra.reduce((n, g) => n + g.series.length, 0);
	const series = chart.series.map((s, i) => seriesView(ctx, s, i, total, collect));
	const longest = Math.max(
		0,
		...series.map((s) => s.values.length),
		...extra.flatMap((g) => g.series.map((s) => s.values.length)),
	);
	const firstCats = rawCategories.find((c) => c.length > 0) ?? [];
	const categories = Array.from({ length: Math.max(longest, firstCats.length) }, (_, i) =>
		firstCats[i] !== undefined ? toText(firstCats[i] ?? null) : String(i + 1),
	);

	if (byPoint) {
		series.forEach((s, seriesIndex) => {
			s.pointColors = s.values.map((_, i) => {
				const explicit = drawingColor(chart.series[seriesIndex]?.pointColors?.[i]);
				return (
					explicit ??
					(palette
						? chartPaletteSeriesColor(palette, i, s.values.length, scheme)
						: autoSeriesColor(theme, i))
				);
			});
		});
	} else
		series.forEach((s, seriesIndex) => {
			const points = chart.series[seriesIndex]?.pointColors;
			if (points) s.pointColors = s.values.map((_, i) => drawingColor(points[i]) ?? s.color);
		});
	series.forEach((s, index) => {
		const source = chart.series[index]!;
		if (s.gradient && !radial && s.pointColors) {
			// Holes inherit the series paint, which becomes a gradient URL at the painter boundary.
			for (let i = 0; i < s.pointColors.length; i++)
				if (!source.pointColors?.[i]) delete s.pointColors[i];
		}
		for (const [key, fill] of Object.entries(source.pointFills ?? {})) {
			if (fill.kind === 'none') (s.pointColors ??= [])[Number(key)] = 'none';
			else if (fill.kind === 'solid') {
				const paint = drawingColor(fill.color);
				if (paint) (s.pointColors ??= [])[Number(key)] = paint;
			}
		}
	});

	const model: ChartViewModel = {
		...(chart.barGapWidth === undefined ? {} : { barGapWidth: chart.barGapWidth }),
		...(chart.barOverlap === undefined ? {} : { barOverlap: chart.barOverlap }),
		type,
		grouping,
		showLegend: chart.showLegend,
		legendPosition: chart.legendPosition ?? 'r',
		categories,
		series,
		horizontal: type === 'bar',
		supported: SUPPORTED.has(type),
	};
	if (chart.title) model.title = chart.title;
	const titleFormatting = chart.formatting?.entries.title;
	if (titleFormatting?.layout) model.titleLayout = titleFormatting.layout;
	if (titleFormatting?.overlay !== undefined) model.titleOverlay = titleFormatting.overlay;
	const legendFormatting = chart.formatting?.entries.legend;
	if (legendFormatting?.layout) model.legendLayout = legendFormatting.layout;
	if (legendFormatting?.overlay !== undefined) model.legendOverlay = legendFormatting.overlay;
	const titleText = chartTitleText(chart, theme);
	if (titleText) model.titleText = titleText;
	const appearance = chartAppearance(chart, theme);
	if (appearance) model.appearance = appearance;

	if (type === 'scatter' || type === 'bubble') {
		series.forEach((s, i) => {
			const cats = rawCategories[i] ?? [];
			const numeric = cats.length > 0 && cats.every((c) => typeof c === 'number');
			s.xValues = s.values.map((_, k) => (numeric ? toNumber(cats[k] ?? null) : k + 1));
		});
		const xs = series.flatMap((s) => s.xValues ?? []).filter((x): x is number => x !== null);
		model.xAxis = xs.length ? niceScale(Math.min(...xs), Math.max(...xs)) : niceScale(0, 1);
	}
	if (type === 'stock') model.stock = stockView(chart.stock, ctx);
	if (type === 'bubble') bubbleView(chart, model, ctx.evaluateRef);
	if (type === 'treemap' || type === 'sunburst') hierarchyView(chart, model, ctx.evaluateRef);
	if (type === 'histogram') histogramView(chart, model);
	if (extra.length && (type === 'column' || type === 'line' || type === 'area'))
		chartExtraView(chart, model, ctx, total, collect);
	if (type === 'waterfall') waterfallView(chart, model, ctx);
	else if (!byPoint) model.valueAxis = valueAxis(model);
	applyDataLabels(model, chart.series, extra);
	chartExtraAxes(model);
	applyChartAxes(model, chart, ctx.drawingColor);
	return model;
}
