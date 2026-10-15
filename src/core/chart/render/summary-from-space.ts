// Builds the painter's chart summary from the format-neutral chart model (`parseChartSpace`): the
// first chart group decides the chart type, its series give the cached data, references and
// colours. Lifted from the spreadsheet chart reader so every format maps a chart part one way.
import { chartSourceValues } from '../data-cache';
import type { ChartGroupKind, ChartSpace } from '../model';
import type { ChartShapeProperties, ChartSpaceSeries } from '../model-series';
import type { ChartStyleDefinition } from '../style-definition';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type {
	ChartSummary,
	ChartSummaryGroup,
	ChartSummarySeries,
	ChartSummaryType,
} from './summary';

const solidColor = (fill: DrawingFill | undefined): DrawingColor | undefined =>
	fill?.kind === 'solid' ? fill.color : undefined;

/** Whether a fill is written as `a:solidFill` (even one whose colour could not be read). */
const isSolidFill = (fill: DrawingFill | undefined): boolean =>
	fill?.kind === 'solid' || (fill?.kind === 'unsupported' && fill.element === 'solidFill');

const GROUP_TYPES: Record<ChartGroupKind, ChartSummaryType> = {
	bar: 'bar',
	line: 'line',
	area: 'area',
	pie: 'pie',
	ofPie: 'pie',
	doughnut: 'doughnut',
	scatter: 'scatter',
	radar: 'radar',
	bubble: 'bubble',
	stock: 'stock',
	surface: 'surface',
};

const GROUPINGS = new Set(['clustered', 'stacked', 'percentStacked', 'standard']);

/** One series of the summary: cached data, references and its primary DrawingML paint. */
export function chartSummarySeries(
	source: ChartSpaceSeries,
	chartType: ChartSummaryType,
): ChartSummarySeries {
	const categories = source.categories ?? source.xValues;
	const values = source.values ?? source.yValues;
	const numeric = (v: string | number | boolean | null | undefined): number | null =>
		typeof v === 'number' ? v : v === '' ? null : Number.isFinite(Number(v)) ? Number(v) : null;
	const out: ChartSummarySeries = {
		categories: chartSourceValues(categories),
		values: chartSourceValues(values).map(numeric),
	};
	if (source.bubbleSizes) {
		out.bubbleSizes = chartSourceValues(source.bubbleSizes).map(numeric);
		if (source.bubbleSizes.formula) out.bubbleSizesRef = source.bubbleSizes.formula;
	}
	if (source.tx?.text !== undefined) out.name = source.tx.text;
	const nameRef = source.tx?.reference?.formula;
	if (nameRef) out.nameRef = nameRef;
	if (categories?.formula) out.categoriesRef = categories.formula;
	if (values?.formula) out.valuesRef = values.formula;
	if (source.spPr?.effectsXml) out.effectsXml = source.spPr.effectsXml;
	// Lines draw in their outline colour, other series in their fill; scatter falls back to markers.
	const primary = ['line', 'scatter', 'radar'].includes(chartType)
		? source.spPr?.line?.fill
		: source.spPr?.fill;
	const fill =
		chartType === 'scatter' && !isSolidFill(primary) ? source.marker?.spPr?.fill : primary;
	const drawingColor = solidColor(fill);
	if (drawingColor) out.drawingColor = drawingColor;
	if (fill && fill.kind !== 'solid' && !(chartType === 'scatter' && fill.kind === 'none'))
		out.fill = fill;
	const points: NonNullable<ChartSummarySeries['pointColors']> = {};
	const pointFills: NonNullable<ChartSummarySeries['pointFills']> = {};
	for (const point of source.dataPoints) {
		const properties: ChartShapeProperties | undefined = point.spPr ?? point.marker?.spPr;
		if (point.index === undefined) continue;
		const pointColor = solidColor(properties?.fill);
		if (pointColor) points[point.index] = pointColor;
		if (properties?.fill && properties.fill.kind !== 'solid')
			pointFills[point.index] = properties.fill;
	}
	if (Object.keys(points).length) out.pointColors = points;
	if (Object.keys(pointFills).length) out.pointFills = pointFills;
	return out;
}

/** The id of the value axis a group plots against (its `c:axId` that names a `c:valAx`). */
function valueAxisId(chartSpace: ChartSpace, axisIds: readonly number[]): number | undefined {
	return axisIds.find((id) =>
		chartSpace.plotArea.axes.some((a) => a.id === id && a.kind === 'val'),
	);
}

/**
 * The groups after the first that a combination chart layers over it. Bar, column, line and area
 * groups are kept (bar groups only over a vertical first group); a group on another value axis is
 * secondary. Pie, scatter, radar, stock and surface groups are not combined.
 */
function extraGroups(chartSpace: ChartSpace, firstType: ChartSummaryType): ChartSummaryGroup[] {
	const [first, ...rest] = chartSpace.plotArea.groups;
	if (!first || !['column', 'line', 'area'].includes(firstType)) return [];
	const primary = valueAxisId(chartSpace, first.axisIds);
	const out: ChartSummaryGroup[] = [];
	for (const group of rest) {
		if (group.is3D || !['bar', 'line', 'area'].includes(group.kind)) continue;
		const type = group.kind === 'bar' ? 'column' : (group.kind as 'line' | 'area');
		const entry: ChartSummaryGroup = {
			chartType: type,
			series: group.series.map((source) => chartSummarySeries(source, type)),
			secondary: valueAxisId(chartSpace, group.axisIds) !== primary,
		};
		if (group.grouping && GROUPINGS.has(group.grouping))
			entry.grouping = group.grouping as NonNullable<ChartSummaryGroup['grouping']>;
		if (type === 'column') {
			if (group.gapWidth !== undefined) entry.barGapWidth = group.gapWidth;
			if (group.overlap !== undefined) entry.barOverlap = group.overlap;
		}
		if (entry.series.length) out.push(entry);
	}
	return out;
}

/** Part-level metadata read beside the chart model (`readChartFormatting`, style parts). */
export interface ChartSummaryExtras {
	formatting?: ChartStyleDefinition;
	colorPalette?: number;
	styleDefinition?: ChartStyleDefinition;
}

/** The painter's summary of a parsed `c:chartSpace`. */
export function chartSummaryFromSpace(
	chartSpace: ChartSpace,
	extras: ChartSummaryExtras = {},
): ChartSummary {
	const group = chartSpace.plotArea.groups[0];
	let chartType: ChartSummaryType = group ? GROUP_TYPES[group.kind] : 'column';
	if (chartType === 'bar' && group?.barDirection !== 'bar') chartType = 'column';
	const series = (group?.series ?? []).map((source) => chartSummarySeries(source, chartType));
	const summary: ChartSummary = {
		chartType,
		series,
		showLegend: chartSpace.legend !== undefined,
	};
	const combined = extraGroups(chartSpace, chartType);
	if (combined.length) summary.extraGroups = combined;
	if (chartType === 'bubble' && group) {
		const sizeRepresents = group.sizeRepresents === 'w' ? 'w' : 'area';
		summary.bubble = {
			sizeRepresents,
			...(group.bubbleScale === undefined ? {} : { bubbleScale: group.bubbleScale }),
			showNegativeBubbles: group.showNegativeBubbles === true,
		};
	}
	const grouping = group?.grouping;
	if (grouping && GROUPINGS.has(grouping))
		summary.grouping = grouping as NonNullable<ChartSummary['grouping']>;
	if (chartType === 'bar' || chartType === 'column') {
		for (const [value, field, fallback, min, max] of [
			[group?.gapWidth, 'barGapWidth', 150, 0, 500],
			[group?.overlap, 'barOverlap', 0, -100, 100],
		] as const) {
			const resolved = value ?? fallback;
			if (Number.isInteger(resolved) && resolved >= min && resolved <= max)
				summary[field] = resolved;
		}
	}
	if (chartSpace.title && chartSpace.autoTitleDeleted !== true) {
		const title = chartSpace.title.text;
		if (title !== undefined) summary.title = title;
		else if (series.length === 1 && series[0]?.name) summary.title = series[0].name;
	}
	const legendPos = chartSpace.legend?.position;
	if (legendPos && ['r', 'l', 't', 'b', 'tr'].includes(legendPos))
		summary.legendPosition = legendPos as NonNullable<ChartSummary['legendPosition']>;
	if (extras.formatting) summary.formatting = extras.formatting;
	if (extras.colorPalette !== undefined) summary.colorPalette = extras.colorPalette;
	if (extras.styleDefinition) summary.styleDefinition = extras.styleDefinition;
	return summary;
}
