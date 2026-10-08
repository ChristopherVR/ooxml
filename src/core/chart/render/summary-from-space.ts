// Builds the painter's chart summary from the format-neutral chart model (`parseChartSpace`): the
// first chart group decides the chart type, its series give the cached data, references and
// colours. Lifted from the spreadsheet chart reader so every format maps a chart part one way.
import { chartSourceValues } from '../data-cache';
import type { ChartGroupKind, ChartSpace } from '../model';
import type { ChartShapeProperties, ChartSpaceSeries } from '../model-series';
import type { ChartStyleDefinition } from '../style-definition';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';
import type { ChartSummary, ChartSummarySeries, ChartSummaryType } from './summary';

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
	const out: ChartSummarySeries = {
		categories: chartSourceValues(categories),
		values: chartSourceValues(values).map((v) =>
			typeof v === 'number' ? v : v === '' ? null : Number.isFinite(Number(v)) ? Number(v) : null,
		),
	};
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
