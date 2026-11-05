/**
 * The two models the save planner compares for a new classic chart part
 * (`chart-new-part.ts`): the chart as the part must show it (`after`, with the
 * rules for doubly stated inputs applied) and what the starting part already
 * holds (the baseline). Every field the baseline leaves out is the planner's
 * to write; a field no planner family claims sends the part to the legacy
 * generator.
 *
 * @module utils/chart-new-part-model
 */
import type { PptxChartData, PptxChartSeries, PptxChartStyle } from '../types';
import { chartModelLabelPositions } from './chart-label-positions';
import type { NewChartAxes } from './chart-new-part-axes';
import { same } from './chart-neutral-save-common';

/**
 * Fields the starting part holds (type, data, title, group settings), and fields no chart part
 * holds (the package paths, the colour style part written beside it, state read on load).
 */
const CARRIED = [
	'chartType',
	'categories',
	'title',
	'grouping',
	'groupingStandard',
	'varyColors',
	'barDirection',
	'barShape',
	'radarStyle',
	'wireframe',
	'surfaceTopView',
	'scatterStyle',
	'barGapWidth',
	'barOverlap',
	'firstSliceAngle',
	'doughnutHoleSize',
	'gapDepth',
	'ofPieOptions',
	'bubbleOptions',
	'date1904',
	'chartPartPath',
	'chartRelationshipId',
	'colorPalette',
	'colorMethod',
	'themeAccentColors',
	'colorStylePartPath',
	'colorStyleOriginalPalette',
	'colorStyleOriginalMethod',
	'chartStyleDefinition',
	'titleTextEditable',
	'dataLabelTextEditable',
	'dataLabelSourceShape',
	'dataLabelTextOnly',
	'titleTextOnly',
	'titleTextOnlyStyles',
	'filteredSeries',
	'filteredSeriesTitle',
	'filteredCategoryTitle',
	'userShapes',
	'userShapesXml',
	'externalData',
	'embeddedWorkbookData',
] as const satisfies readonly (keyof PptxChartData)[];

/** Style fields the starting part holds; `hasGridlines` is read on load and never written. */
const CARRIED_STYLE = [
	'hasLegend',
	'legendPosition',
	'hasGridlines',
] as const satisfies readonly (keyof PptxChartStyle)[];

/** Series fields the starting part holds, and the extended-chart options a classic part ignores. */
const CARRIED_SERIES = [
	'name',
	'values',
	'blanks',
	'idx',
	'uniqueId',
	'boxWhiskerOptions',
	'histogramOptions',
	'waterfallOptions',
	'regionMapOptions',
	'treemapOptions',
] as const satisfies readonly (keyof PptxChartSeries)[];

function pick<T extends object>(source: T, keys: readonly (keyof T)[]): Partial<T> {
	const out: Partial<T> = {};
	for (const key of keys)
		if (source[key] !== undefined) out[key] = structuredClone(source[key]) as T[keyof T];
	return out;
}

const HEX = /^#?[\da-f]{6}$/iu;
const LINE_DRAWN = new Set<string>(['line', 'stock']);

/**
 * Label positions the chart type does not offer are left out (and reported): PowerPoint repairs a
 * part with one, as it did for the legacy generator's chart-wide labels.
 */
function dropRefusedLabelPositions(after: PptxChartData, unwritten: string[]): void {
	const refuse = (allowed: readonly string[] | undefined, position: string | undefined) =>
		position !== undefined && allowed !== undefined && !allowed.includes(position);
	const chartWide = after.style?.dataLabels;
	if (refuse(chartModelLabelPositions(after), chartWide?.position)) {
		delete chartWide!.position;
		unwritten.push('style.dataLabels.position');
	}
	after.series.forEach((series, index) => {
		const allowed = chartModelLabelPositions(after, index);
		if (refuse(allowed, series.dataLabelOptions?.position)) {
			delete series.dataLabelOptions!.position;
			unwritten.push(`series.${index}.dataLabelOptions.position`);
		}
		series.dataLabels?.forEach((label, point) => {
			if (!refuse(allowed, label.position)) return;
			delete label.position;
			unwritten.push(`series.${index}.dataLabels.${point}.position`);
		});
	});
}

/** `chart` as the new part must show it, with the fields left out of it as paths. */
export function normalizedNewChart(
	chart: PptxChartData,
	options: { labelPositionsOnly?: boolean } = {},
): { after: PptxChartData; unwritten: string[] } {
	const after = structuredClone(chart);
	const unwritten: string[] = [];
	dropRefusedLabelPositions(after, unwritten);
	if (options.labelPositionsOnly) return { after, unwritten };
	// The generator wrote hex colours upper-case (`a:srgbClr/@val`).
	for (const series of after.series)
		if (series.color && HEX.test(series.color)) series.color = series.color.toUpperCase();
	// A 2-D line or stock series with no smoothing of its own is straight: PowerPoint smooths one
	// that writes no `c:smooth`, so the part says so (a 3-D line has no smoothing).
	for (const series of after.series)
		if (LINE_DRAWN.has(series.seriesChartType ?? after.chartType)) series.smooth ??= false;
	// A scatter series without its own smoothing follows the style: smoothed for `smoothMarker`.
	if (after.chartType === 'scatter')
		for (const series of after.series)
			series.smooth ??= (after.scatterStyle ?? 'lineMarker') === 'smoothMarker';
	// The documented x values and bubble sizes of a new scatter or bubble chart without its own.
	if (after.chartType === 'scatter' || after.chartType === 'bubble') {
		const xs = after.categories.map((category, index) => {
			const number = Number.parseFloat(category);
			return Number.isFinite(number) ? number : index + 1;
		});
		for (const series of after.series) {
			series.xValues ??= [...xs];
			if (after.chartType === 'bubble') series.bubbleSizes ??= series.values.map(() => 1);
		}
	}
	if (after.style?.hasTitle === false) {
		delete after.title;
		delete after.titleRuns;
	} else if (after.titleRuns?.length) {
		// The runs are the title; the planner's run family takes them from their first run's text.
		after.title = after.titleRuns[0]!.text;
	} else if (after.title === '') {
		// An empty title is no title, as the generator always treated it.
		delete after.title;
	}
	return { after, unwritten };
}

/** What the starting part of `after` holds, in the model's terms. */
export function newChartBaseline(after: PptxChartData, axes: NewChartAxes): PptxChartData {
	const levels = after.categoryLevels;
	const flatLevels =
		levels !== undefined && levels.length <= 1 && same(levels[0] ?? [], after.categories);
	const plotted = new Map(axes.secondary.map((item) => [item.seriesIndex, item]));
	// The second slot is the value axis (the Y value axis of a scatter or bubble chart).
	const primaryValueAxis = axes.after[1];
	const style = after.style ? pick(after.style, CARRIED_STYLE) : undefined;
	return {
		...pick(after, CARRIED),
		chartType: after.chartType,
		categories: [...after.categories],
		...(flatLevels ? { categoryLevels: structuredClone(levels) } : {}),
		...(after.chartType === 'stock' && after.hiLowLines === null ? { hiLowLines: null } : {}),
		series: after.series.map((series, index) => {
			const own = pick(series, CARRIED_SERIES) as PptxChartSeries;
			const onPrimary = series.axisId !== undefined && series.axisId === primaryValueAxis?.axisId;
			if (plotted.has(index) || onPrimary) own.axisId = series.axisId!;
			return own;
		}),
		...(style && Object.keys(style).length > 0 ? { style } : {}),
		...(after.axes !== undefined ? { axes: structuredClone(axes.baseline) } : {}),
	};
}
