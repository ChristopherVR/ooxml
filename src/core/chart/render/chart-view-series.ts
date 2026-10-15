// Resolving one summary series into its painter view: live or cached values, its name and the
// colour chosen from its DrawingML paint, the host colour, the colour style or the theme. Shared by
// the first chart group and the groups a combination chart layers over it.
import { chartPaletteSeriesColor, findChartColorPalette } from '../color-palettes';
import { resolveDrawingColor } from '../../drawingml/drawing-color';
import { drawingColorCss } from '../../drawingml/drawing-color-css';
import { resolveDrawingShadowXml } from '../../drawingml/drawing-shadow';
import { resolveChartGradient, type ChartGradientFill } from '../gradient-definition';
import { autoSeriesColor, chartColorScheme } from './chart-colors';
import type { ChartSeriesView } from './chart-view';
import type { ChartSummary, ChartSummarySeries, ChartThemePalette } from './summary';

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

export function toText(value: ChartRefValue): string {
	if (value === null) return '';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (isRefError(value)) return value.error;
	return String(value);
}

export const toNumber = (value: ChartRefValue): number | null =>
	typeof value === 'number' && Number.isFinite(value) ? value : null;

export function resolveRef(
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

/** What resolving a chart's series needs: the theme, colour style and reference evaluator. */
export interface SeriesContext<S extends ChartSummarySeries> {
	theme: ChartThemePalette;
	scheme: Readonly<Record<string, string>>;
	palette: ReturnType<typeof findChartColorPalette>;
	evaluateRef: EvaluateRef;
	seriesColor: ChartViewOptions<S>['seriesColor'];
	drawingColor: (color: ChartSummarySeries['drawingColor']) => string | undefined;
	gradient: (fill: ChartSummarySeries['fill']) => ChartGradientFill | undefined;
}

export function seriesContext<S extends ChartSummarySeries>(
	chart: Pick<ChartSummary<S>, 'colorPalette'>,
	theme: ChartThemePalette,
	options: ChartViewOptions<S>,
): SeriesContext<S> {
	const scheme = chartColorScheme(theme) as Readonly<Record<string, string>>;
	const resolver = { scheme: (name: string) => scheme[name] };
	return {
		theme,
		scheme,
		palette:
			chart.colorPalette === undefined ? undefined : findChartColorPalette(chart.colorPalette),
		evaluateRef: options.evaluateRef ?? (() => []),
		seriesColor: options.seriesColor,
		drawingColor: (color) =>
			color &&
			drawingColorCss(resolveDrawingColor(color, resolver, { transformOrder: 'document' })),
		gradient: (fill) =>
			fill?.kind === 'gradient'
				? resolveChartGradient(fill, (color) =>
						resolveDrawingColor(color, resolver, { transformOrder: 'document' }),
					)
				: undefined,
	};
}

/**
 * One series' view. `index` is its position among all series of the chart (it picks the automatic
 * colour) and `total` the count of all of them (the colour style spreads over it).
 */
export function seriesView<S extends ChartSummarySeries>(
	ctx: SeriesContext<S>,
	s: ChartSummarySeries,
	index: number,
	total: number,
	categoriesOf: (cats: ChartRefValue[]) => void,
): ChartSeriesView {
	const { theme, scheme, palette, evaluateRef } = ctx;
	const values = (resolveRef(s.valuesRef, evaluateRef)?.map(toNumber) ??
		s.values.map(toNumber)) as (number | null)[];
	categoriesOf(resolveRef(s.categoriesRef, evaluateRef) ?? s.categories);
	const paletteColor = palette && chartPaletteSeriesColor(palette, index, total, scheme);
	const color =
		s.fill?.kind === 'none'
			? 'none'
			: (ctx.drawingColor(s.fill?.kind === 'solid' ? s.fill.color : s.drawingColor) ??
				ctx.seriesColor?.(s as S, index) ??
				paletteColor ??
				autoSeriesColor(theme, index));
	const view: ChartSeriesView = { name: seriesName(s, index, evaluateRef), values, color };
	const shadow = resolveDrawingShadowXml(s.effectsXml, { scheme: (name) => scheme[name] });
	if (shadow) view.shadow = shadow;
	const resolvedGradient = ctx.gradient(s.fill);
	if (resolvedGradient) view.gradient = resolvedGradient;
	const points = Object.fromEntries(
		Object.entries(s.pointFills ?? {}).flatMap(([key, fill]) => {
			const value = ctx.gradient(fill);
			return value ? [[key, value]] : [];
		}),
	);
	if (Object.keys(points).length) view.pointGradients = points;
	return view;
}
