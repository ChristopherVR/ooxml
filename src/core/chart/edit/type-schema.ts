// What each chart group element and series type may hold, in model field names (ECMA-376 Part 1,
// 21.2.2: CT_BarChart, CT_LineChart... and CT_BarSer, CT_LineSer...). The same content models as
// PowerPoint's `chart-container-content-model.ts`, expressed over the neutral model so a type switch
// keeps exactly what the new type allows.
import type { ChartGroupKind, ChartPlotGroup } from '../model';
import type { ChartDataPoint, ChartSpaceSeries } from '../model-series';
import { isXYChartKind } from './series';

type GroupField = keyof ChartPlotGroup;

const DROP_LINES: GroupField[] = ['dropLines', 'dropLinesSpPr'];
const SERIES_LINES: GroupField[] = ['seriesLines', 'seriesLinesSpPr'];
const LABELLED: GroupField[] = ['varyColors', 'dataLabels', 'extLst'];

/** The optional settings each chart group element allows, beside its series and axis ids. */
export const GROUP_FIELDS: Readonly<Record<string, readonly GroupField[]>> = {
	barChart: [...LABELLED, 'barDirection', 'grouping', 'gapWidth', 'overlap', ...SERIES_LINES],
	bar3DChart: [...LABELLED, 'barDirection', 'grouping', 'gapWidth', 'gapDepth', 'shape'],
	lineChart: [
		...LABELLED,
		'grouping',
		...DROP_LINES,
		'hiLowLines',
		'hiLowLinesSpPr',
		'upDownBars',
		'upDownBarsGapWidth',
		'upBars',
		'downBars',
		'marker',
		'smooth',
	],
	line3DChart: [...LABELLED, 'grouping', ...DROP_LINES, 'gapDepth'],
	areaChart: [...LABELLED, 'grouping', ...DROP_LINES],
	area3DChart: [...LABELLED, 'grouping', ...DROP_LINES, 'gapDepth'],
	pieChart: [...LABELLED, 'firstSliceAngle'],
	pie3DChart: [...LABELLED],
	doughnutChart: [...LABELLED, 'firstSliceAngle', 'holeSize'],
	ofPieChart: [
		...LABELLED,
		'ofPieType',
		'gapWidth',
		'splitType',
		'splitPosition',
		'customSplit',
		'secondPieSize',
		...SERIES_LINES,
	],
	scatterChart: [...LABELLED, 'scatterStyle'],
	radarChart: [...LABELLED, 'radarStyle'],
	bubbleChart: [...LABELLED, 'bubble3D', 'bubbleScale', 'showNegativeBubbles', 'sizeRepresents'],
	stockChart: ['dataLabels', 'extLst', ...DROP_LINES, 'hiLowLines', 'hiLowLinesSpPr', 'upDownBars'],
	surfaceChart: ['wireframe', 'bandFormatsXml', 'extLst'],
	surface3DChart: ['wireframe', 'bandFormatsXml', 'extLst'],
};

/** Whether a group element references axes (the pie family does not). */
export const groupHasAxes = (element: string): boolean =>
	!['pieChart', 'pie3DChart', 'doughnutChart', 'ofPieChart'].includes(element);

type SeriesField = keyof ChartSpaceSeries;
const ANALYSED: SeriesField[] = ['dataPoints', 'dataLabels', 'trendlines', 'errorBars'];

/** The optional series fields of each family, beside idx, order, tx, spPr and extLst. */
const SERIES_FIELDS: Readonly<Record<ChartGroupKind, readonly SeriesField[]>> = {
	bar: ['invertIfNegative', 'pictureOptionsXml', ...ANALYSED, 'categories', 'values', 'shape'],
	line: ['marker', ...ANALYSED, 'categories', 'values', 'smooth'],
	stock: ['marker', ...ANALYSED, 'categories', 'values', 'smooth'],
	area: ['pictureOptionsXml', ...ANALYSED, 'categories', 'values'],
	scatter: ['marker', ...ANALYSED, 'xValues', 'yValues', 'smooth'],
	bubble: ['invertIfNegative', ...ANALYSED, 'xValues', 'yValues', 'bubbleSizes', 'bubble3D'],
	pie: ['explosion', 'dataPoints', 'dataLabels', 'categories', 'values'],
	doughnut: ['explosion', 'dataPoints', 'dataLabels', 'categories', 'values'],
	ofPie: ['explosion', 'dataPoints', 'dataLabels', 'categories', 'values'],
	radar: ['marker', 'dataPoints', 'dataLabels', 'categories', 'values'],
	surface: ['categories', 'values'],
};

/** The optional data point fields of each family (CT_DPt is shared, its meaning is not). */
function pointFor(point: ChartDataPoint, kind: ChartGroupKind): ChartDataPoint {
	const next = { ...point };
	if (kind !== 'bar' && kind !== 'bubble') delete next.invertIfNegative;
	if (!['line', 'stock', 'scatter', 'radar'].includes(kind)) delete next.marker;
	if (kind !== 'bubble') delete next.bubble3D;
	if (kind !== 'pie' && kind !== 'doughnut' && kind !== 'ofPie') delete next.explosion;
	return next;
}

const ALWAYS: SeriesField[] = ['index', 'order', 'tx', 'spPr', 'extLst'];

/**
 * A series of a `from` group as a series of a `to` group: categories and values become x and y
 * values for scatter and bubble charts and back, and fields the new type does not allow are
 * dropped (a category chart keeps one set of y error bars).
 */
export function stripSeriesForKind(
	series: ChartSpaceSeries,
	from: ChartGroupKind,
	to: ChartGroupKind,
): ChartSpaceSeries {
	const source: ChartSpaceSeries = { ...series };
	if (isXYChartKind(to) && !isXYChartKind(from)) {
		if (source.categories) source.xValues = source.categories;
		if (source.values) source.yValues = source.values;
	} else if (!isXYChartKind(to) && isXYChartKind(from)) {
		if (source.xValues) source.categories = source.xValues;
		if (source.yValues) source.values = source.yValues;
	}
	const allowed = new Set<SeriesField>([...ALWAYS, ...SERIES_FIELDS[to]]);
	const next = { dataPoints: [] } as unknown as ChartSpaceSeries;
	for (const key of Object.keys(source) as SeriesField[])
		if (allowed.has(key) && source[key] !== undefined) Object.assign(next, { [key]: source[key] });
	next.dataPoints = (next.dataPoints ?? []).map((point) => pointFor(point, to));
	if (next.errorBars && !isXYChartKind(to)) {
		const bars = next.errorBars.filter((bar) => bar.direction !== 'x').slice(0, 1);
		if (bars.length) next.errorBars = bars;
		else delete next.errorBars;
	}
	return next;
}
