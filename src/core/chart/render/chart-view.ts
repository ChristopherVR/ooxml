import type {
	ChartSummary,
	ChartSummarySeries,
	ChartSummaryType,
	ChartThemePalette,
} from './summary';
import { autoSeriesColor, chartColorScheme } from './chart-colors';
import { chartPaletteSeriesColor, findChartColorPalette } from '../color-palettes';
import { resolveDrawingColor } from '../../drawingml/drawing-color';
import { drawingColorCss } from '../../drawingml/drawing-color-css';
import { niceScale, type AxisScale } from './chart-scale';
import { valueAxis } from './chart-value-axis';
export { categoryTotals } from './chart-value-axis';
import { chartAppearance, type ChartAppearance } from './chart-appearance';
import { resolveChartGradient, type ChartGradientFill } from '../gradient-definition';
import { resolveDrawingShadowXml, type DrawingSvgShadow } from '../../drawingml/drawing-shadow';
import { chartTitleText, type ChartTitleText } from './chart-title-text';
import type { ChartManualLayout } from '../manual-layout';

export interface ChartSeriesView {
	shadow?: DrawingSvgShadow;
	shadowFilter?: string;
	gradient?: ChartGradientFill;
	pointGradients?: Record<number, ChartGradientFill>;
	name: string;
	values: (number | null)[];
	/** Scatter X values (numeric categories, or 1..n when the categories are text). */
	xValues?: (number | null)[];
	color: string;
	/** Per-point colours (pie and doughnut vary colours by point). */
	pointColors?: string[];
}

export interface ValueAxisView extends AxisScale {
	/** Ticks are fractions shown as percentages (percent-stacked charts). */
	percent: boolean;
}

/** Neutral data for a chart painter: series resolved, colours chosen, axes scaled. */
export interface ChartViewModel {
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
	/** Scatter X axis. */
	xAxis?: AxisScale;
	/** Bars run horizontally (`bar`), not vertically (`column`). */
	horizontal: boolean;
	/** False for chart types this painter does not draw (bubble, stock, surface). */
	supported: boolean;
}

/** A value a series reference resolves to: a number, text, boolean, blank or an error value. */
export type ChartRefValue = number | string | boolean | null | { readonly error: string };

export type EvaluateRef = (ref: string) => ChartRefValue[];

/** How a host resolves live references and its own legacy series colours. */
export interface ChartViewOptions<S extends ChartSummarySeries> {
	/** Live values of a series reference; cached values are used when absent or empty. */
	evaluateRef?: EvaluateRef;
	/** A host colour for the series, used after its DrawingML paint and before the palette. */
	seriesColor?: (series: S, index: number) => string | undefined;
}

const isRefError = (value: unknown): value is { readonly error: string } =>
	typeof value === 'object' && value !== null && 'error' in value;

const SUPPORTED: ReadonlySet<ChartSummaryType> = new Set([
	'bar',
	'column',
	'line',
	'area',
	'pie',
	'doughnut',
	'scatter',
	'radar',
]);

function toText(value: ChartRefValue): string {
	if (value === null) return '';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (isRefError(value)) return value.error;
	return String(value);
}

const toNumber = (value: ChartRefValue): number | null =>
	typeof value === 'number' && Number.isFinite(value) ? value : null;

function resolveRef(
	ref: string | undefined,
	evaluateRef: EvaluateRef,
): ChartRefValue[] | undefined {
	if (!ref) return undefined;
	try {
		const values = evaluateRef(ref);
		return values.length ? values : undefined;
	} catch {
		return undefined;
	}
}

function seriesName(series: ChartSummarySeries, index: number, evaluateRef: EvaluateRef): string {
	const fromRef = resolveRef(series.nameRef, evaluateRef)
		?.map(toText)
		.filter((t) => t !== '')
		.join(' ');
	return fromRef || series.name || `Series${index + 1}`;
}

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
	const evaluateRef: EvaluateRef = options.evaluateRef ?? (() => []);
	const type = chart.chartType;
	const radial = type === 'pie' || type === 'doughnut';
	const palette =
		chart.colorPalette === undefined ? undefined : findChartColorPalette(chart.colorPalette);
	const scheme = chartColorScheme(theme);
	const drawingColor = (color: ChartSummarySeries['drawingColor']) =>
		color &&
		drawingColorCss(
			resolveDrawingColor(
				color,
				{
					scheme: (name) => (scheme as Readonly<Record<string, string>>)[name],
				},
				{ transformOrder: 'document' },
			),
		);
	const gradient = (fill: ChartSummarySeries['fill']) =>
		fill?.kind === 'gradient'
			? resolveChartGradient(fill, (color) =>
					resolveDrawingColor(
						color,
						{
							scheme: (name) => (scheme as Readonly<Record<string, string>>)[name],
						},
						{ transformOrder: 'document' },
					),
				)
			: undefined;
	const grouping =
		chart.grouping ?? (type === 'bar' || type === 'column' ? 'clustered' : 'standard');
	let categories: string[] = [];
	const rawCategories: ChartRefValue[][] = [];
	const series: ChartSeriesView[] = chart.series.map((s, i) => {
		const values = (resolveRef(s.valuesRef, evaluateRef)?.map(toNumber) ??
			s.values.map(toNumber)) as (number | null)[];
		const cats = resolveRef(s.categoriesRef, evaluateRef) ?? s.categories;
		rawCategories.push(cats);
		const paletteColor =
			palette && chartPaletteSeriesColor(palette, i, chart.series.length, scheme);
		const color =
			s.fill?.kind === 'none'
				? 'none'
				: (drawingColor(s.fill?.kind === 'solid' ? s.fill.color : s.drawingColor) ??
					options.seriesColor?.(s, i) ??
					paletteColor ??
					autoSeriesColor(theme, i));
		const view: ChartSeriesView = { name: seriesName(s, i, evaluateRef), values, color };
		const shadow = resolveDrawingShadowXml(s.effectsXml, {
			scheme: (name) => (scheme as Readonly<Record<string, string>>)[name],
		});
		if (shadow) view.shadow = shadow;
		const resolvedGradient = gradient(s.fill);
		if (resolvedGradient) view.gradient = resolvedGradient;
		const points = Object.fromEntries(
			Object.entries(s.pointFills ?? {}).flatMap(([key, fill]) => {
				const value = gradient(fill);
				return value ? [[key, value]] : [];
			}),
		);
		if (Object.keys(points).length) view.pointGradients = points;
		return view;
	});
	const longest = Math.max(0, ...series.map((s) => s.values.length));
	const firstCats = rawCategories.find((c) => c.length > 0) ?? [];
	categories = Array.from({ length: Math.max(longest, firstCats.length) }, (_, i) =>
		firstCats[i] !== undefined ? toText(firstCats[i] ?? null) : String(i + 1),
	);

	if (radial) {
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

	if (type === 'scatter') {
		series.forEach((s, i) => {
			const cats = rawCategories[i] ?? [];
			const numeric = cats.length > 0 && cats.every((c) => typeof c === 'number');
			s.xValues = s.values.map((_, k) => (numeric ? toNumber(cats[k] ?? null) : k + 1));
		});
		const xs = series.flatMap((s) => s.xValues ?? []).filter((x): x is number => x !== null);
		model.xAxis = xs.length ? niceScale(Math.min(...xs), Math.max(...xs)) : niceScale(0, 1);
	}
	if (!radial) model.valueAxis = valueAxis(model);
	return model;
}
