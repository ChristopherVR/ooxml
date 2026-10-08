// Builds the spreadsheet chart object from the format-neutral chart model (`chart/parseChartSpace`):
// the first chart group decides the chart type, its series give the cached data, references and
// colours. Direct formatting metadata is still read from the part by `chart/readChartFormatting`.
import { parseXml } from '../../xml/index';
import type { ChartObject, ChartSeries, ChartType, Color, DrawingAnchor } from '../model';
import { readChartFormatting } from '../../chart/read-formatting';
import { chartSourceValues } from '../../chart/data-cache';
import { parseChartSpace } from '../../chart/parse-space';
import type { ChartGroupKind } from '../../chart/model';
import type { ChartShapeProperties, ChartSpaceSeries } from '../../chart/model-series';
import type { DrawingColor, DrawingFill } from '../../drawingml/types';

/** DrawingML scheme colour names to SpreadsheetML theme indices. */
const SCHEME_INDEX: Record<string, number> = {
	bg1: 0,
	lt1: 0,
	tx1: 1,
	dk1: 1,
	bg2: 2,
	lt2: 2,
	tx2: 3,
	dk2: 3,
	accent1: 4,
	accent2: 5,
	accent3: 6,
	accent4: 7,
	accent5: 8,
	accent6: 9,
	hlink: 10,
	folHlink: 11,
};

/** The SpreadsheetML colour of a plain solid fill colour (RGB or a scheme slot), if it has one. */
function legacyColor(color: DrawingColor | undefined): Color | undefined {
	if (color?.kind === 'srgb') return color.value ? { rgb: color.value.toUpperCase() } : undefined;
	if (color?.kind === 'scheme') {
		const index = SCHEME_INDEX[color.value];
		return index === undefined ? undefined : { theme: index };
	}
	return undefined;
}

const solidColor = (fill: DrawingFill | undefined): DrawingColor | undefined =>
	fill?.kind === 'solid' ? fill.color : undefined;

/** Whether a fill is written as `a:solidFill` (even one whose colour could not be read). */
const isSolidFill = (fill: DrawingFill | undefined): boolean =>
	fill?.kind === 'solid' || (fill?.kind === 'unsupported' && fill.element === 'solidFill');

const GROUP_TYPES: Record<ChartGroupKind, ChartType> = {
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

function toSeries(source: ChartSpaceSeries, chartType: ChartType): ChartSeries {
	const categories = source.categories ?? source.xValues;
	const values = source.values ?? source.yValues;
	const out: ChartSeries = {
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
	const color = legacyColor(solidColor(fill));
	if (color) out.color = color;
	const drawingColor = solidColor(fill);
	if (drawingColor) out.drawingColor = drawingColor;
	if (fill && fill.kind !== 'solid' && !(chartType === 'scatter' && fill.kind === 'none'))
		out.fill = fill;
	const points: NonNullable<ChartSeries['pointColors']> = {};
	const pointFills: NonNullable<ChartSeries['pointFills']> = {};
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

/** Reads the modelled summary of a `c:chartSpace` part. */
export function parseChart(
	xml: string,
	anchor: DrawingAnchor,
	partName: string,
	name?: string,
): ChartObject {
	const root = parseXml(xml, { label: 'XLSX chart' }).documentElement;
	const { chartSpace } = parseChartSpace(root);
	const group = chartSpace.plotArea.groups[0];
	let chartType: ChartType = group ? GROUP_TYPES[group.kind] : 'column';
	if (chartType === 'bar' && group?.barDirection !== 'bar') chartType = 'column';
	const series = (group?.series ?? []).map((source) => toSeries(source, chartType));
	const object: ChartObject = {
		kind: 'chart',
		anchor,
		chartType,
		series,
		showLegend: chartSpace.legend !== undefined,
		partName,
	};
	const grouping = group?.grouping;
	if (grouping && GROUPINGS.has(grouping))
		object.grouping = grouping as NonNullable<ChartObject['grouping']>;
	if (chartType === 'bar' || chartType === 'column') {
		for (const [value, field, fallback, min, max] of [
			[group?.gapWidth, 'barGapWidth', 150, 0, 500],
			[group?.overlap, 'barOverlap', 0, -100, 100],
		] as const) {
			const resolved = value ?? fallback;
			if (Number.isInteger(resolved) && resolved >= min && resolved <= max)
				object[field] = resolved;
		}
	}
	if (chartSpace.title && chartSpace.autoTitleDeleted !== true) {
		const title = chartSpace.title.text;
		if (title !== undefined) object.title = title;
		else if (series.length === 1 && series[0]?.name) object.title = series[0].name;
	}
	const legendPos = chartSpace.legend?.position;
	if (legendPos && ['r', 'l', 't', 'b', 'tr'].includes(legendPos))
		object.legendPosition = legendPos as NonNullable<ChartObject['legendPosition']>;
	if (name) object.name = name;
	const formatting = readChartFormatting(root);
	if (formatting) object.formatting = formatting;
	return object;
}
