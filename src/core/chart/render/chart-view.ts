import { resolveRegionValueColors } from '../region-map/value-colors';
import { chartDataTableView, type ChartDataTableView } from './chart-view-data-table';
import type { ChartMarkerSeriesView } from './chart-view-markers';
import { applyChartAxes, type ChartAxesView } from './chart-view-axes';
import { extentOf } from '../extent';
import type { ParentLabelLayout } from '../hierarchy';
import { bubbleView, hierarchyView } from './chart-view-families';
import { stockView, type StockView } from './chart-view-stock';
import { resolveSurfaces3D, type Chart3DSurfacesView } from './chart-3d-walls';
import type {
	BubbleSummaryOptions,
	Chart3DBarShape,
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
import type { ChartPicturePaint } from './chart-picture-paint';
import type { DrawingSvgShadow } from '../../drawingml/drawing-shadow';
import { chartTitleText, type ChartTitleText } from './chart-title-text';
import type { ChartErrorBarsView, ChartTrendlineView } from './chart-view-analysis';
import type { ChartManualLayout } from '../manual-layout';
import { applyDataLabels, type SeriesLabelViews } from './chart-view-labels';
import {
	chartExtraAxes,
	chartExtraView,
	histogramView,
	paretoAxis,
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

export interface ChartSeriesView extends ChartMarkerSeriesView {
	boxWhisker?: import('../ex/model').ChartExLayoutProperties;
	boxCategories?: string[];
	hidden?: boolean;
	numberFormat?: string;
	smooth?: boolean;
	showLine?: boolean;
	stroke?: import('./chart-view-analysis').ChartStrokeView;
	/** Document-order position, retained when combination groups are painted together. */
	sourceIndex?: number;
	shadow?: DrawingSvgShadow;
	shadowFilter?: string;
	gradient?: ChartGradientFill;
	pointGradients?: Record<number, ChartGradientFill>;
	/** A picture fill (`a:blipFill`) of the series, and of individual points. */
	picture?: ChartPicturePaint;
	pointPictures?: Record<number, ChartPicturePaint>;
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
	/** 3-D bar shape override on the series (`c:ser/c:shape`). */
	barShape?: Chart3DBarShape;
	/** Pie slice explosion percent (`c:ser/c:explosion`). */
	explosion?: number;
	/** Per-point pie slice explosion percents (`c:dPt/c:explosion`). */
	pointExplosions?: Record<number, number>;
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
	funnelGapWidth?: number;
	regionMap?: import('../region-map/types').RegionMapOptions;
	dataTable?: ChartDataTableView;
	dateCategories?: (number | null)[];
	date1904?: boolean;
	frameSize?: import('../manual-layout').ChartFrameSize;
	axes?: ChartAxesView;
	plotLayout?: ChartManualLayout;
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
	/** 3-D charts: the view the painter projects (bar3D, line3D, area3D, pie3D, surface3D). */
	view3D?: import('./chart-3d-projection').Chart3DViewInput;
	/** 3-D bar shape (`c:bar3DChart/c:shape`). */
	barShape?: Chart3DBarShape;
	/** 3-D floor and wall formatting (`c:floor`, `c:sideWall`, `c:backWall`). */
	surfaces3D?: Chart3DSurfacesView;
	/** Surface charts: drawn as a wireframe without filled bands (`c:wireframe`). */
	wireframe?: boolean;
	/** 2-D contour surface chart (`c:surfaceChart`, top-down view). */
	surfaceContour?: boolean;
	/** Pie and doughnut: angle of the first slice in degrees clockwise from 12 o'clock. */
	firstSliceAngle?: number;
	/** False for chart types this painter does not draw (a surface without a 3-D view). */
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
	'funnel',
	'regionMap',
	'boxWhisker',
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
		series: type === 'boxWhisker' ? series.filter((s) => !s.hidden) : series,
		horizontal: type === 'bar',
		supported:
			SUPPORTED.has(type) ||
			(type === 'surface' && (chart.view3D !== undefined || chart.surfaceContour === true)),
	};
	if (chart.view3D) model.view3D = chart.view3D;
	if (chart.barShape) model.barShape = chart.barShape;
	if (chart.wireframe !== undefined) model.wireframe = chart.wireframe;
	if (chart.surfaceContour !== undefined) model.surfaceContour = chart.surfaceContour;
	if (chart.firstSliceAngle !== undefined) model.firstSliceAngle = chart.firstSliceAngle;
	const surfaces3D = resolveSurfaces3D(chart.surfaces3D, ctx.drawingColor, ctx.gradient);
	if (surfaces3D) model.surfaces3D = surfaces3D;
	if (chart.funnelGapWidth !== undefined) model.funnelGapWidth = chart.funnelGapWidth;
	if (chart.series[0]?.regionMap)
		model.regionMap = resolveRegionValueColors(chart.series[0].regionMap, ctx.drawingColor);
	if (model.regionMap && chart.series[0]?.fill?.kind === 'solid' && series[0])
		model.regionMap.fixedColor = series[0].color;
	const table = chartDataTableView(chart, theme, ctx.drawingColor);
	if (table) model.dataTable = table;
	if (chart.date1904 !== undefined) model.date1904 = chart.date1904;
	if (chart.axes?.some((a) => a.kind === 'date')) model.dateCategories = firstCats.map(toNumber);
	if (chart.plotLayout) model.plotLayout = chart.plotLayout;
	if (type === 'boxWhisker')
		model.categories = [...new Set(model.series.flatMap((s) => s.boxCategories ?? []))];
	if (chart.title) model.title = chart.title;
	const titleFormatting = chart.formatting?.entries.title;
	if (titleFormatting?.layout) model.titleLayout = titleFormatting.layout;
	if (titleFormatting?.overlay !== undefined) model.titleOverlay = titleFormatting.overlay;
	const legendFormatting = chart.formatting?.entries.legend;
	if (legendFormatting?.layout) model.legendLayout = legendFormatting.layout;
	if (legendFormatting?.overlay !== undefined) model.legendOverlay = legendFormatting.overlay;
	const titleText = chartTitleText(chart, theme);
	if (titleText) model.titleText = titleText;
	const appearance = chartAppearance(chart, theme, (fill) => ctx.picture(fill));
	if (appearance) model.appearance = appearance;

	if (type === 'scatter' || type === 'bubble') {
		series.forEach((s, i) => {
			const cats = rawCategories[i] ?? [];
			const numeric = cats.length > 0 && cats.every((c) => typeof c === 'number');
			s.xValues = s.values.map((_, k) => (numeric ? toNumber(cats[k] ?? null) : k + 1));
		});
		const xs = series.flatMap((s) => s.xValues ?? []).filter((x): x is number => x !== null);
		const range = extentOf(xs);
		model.xAxis = xs.length ? niceScale(range.min, range.max) : niceScale(0, 1);
	}
	if (type === 'funnel' || type === 'boxWhisker')
		series.forEach((s, i) => {
			const points = chart.series[i]?.pointColors;
			if (points) s.pointColors = s.values.map((_, j) => drawingColor(points[j]) ?? s.color);
		});
	if (type === 'stock') model.stock = stockView(chart.stock, ctx);
	if (type === 'bubble') bubbleView(chart, model, ctx.evaluateRef);
	if (type === 'treemap' || type === 'sunburst') hierarchyView(chart, model, ctx.evaluateRef);
	if (type === 'histogram') histogramView(chart, model, ctx);
	if (extra.length && (type === 'column' || type === 'line' || type === 'area'))
		chartExtraView(chart, model, ctx, total, collect);
	if (type === 'waterfall') waterfallView(chart, model, ctx);
	else if (!byPoint) model.valueAxis = valueAxis(model);
	applyDataLabels(model, chart.series, extra, ctx.drawingColor);
	chartExtraAxes(model);
	if (type === 'histogram' && chart.histogram?.pareto) paretoAxis(model);
	applyChartAxes(model, chart, ctx.drawingColor);
	return model;
}
