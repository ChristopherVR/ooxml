// The neutral chart model (`chart/ChartSpace`) of a chart created in the spreadsheet model: what
// `chart/writeChartSpace` serialises for a new chart part. It describes the part the xlsx writer
// has always generated (and Excel accepts): one chart group, two axes with fixed ids (the
// `chart/edit` defaults the type switch also uses), palette or manual series colours, the modern
// built-in style 102 unless the chart carries its own.
import { builtInChartStyleXml, effectiveBuiltInChartStyle } from '../../chart/built-in-text-style';
import { applyChartQuickLayout, chartSupportsQuickLayout } from '../../chart/edit/quick-layout';
import { DEFAULT_CHART_AXIS_IDS, defaultChartAxes } from '../../chart/edit/type-axes';
import type { ChartAxis, ChartPlotGroup, ChartSpace } from '../../chart/model';
import type {
	ChartDataPoint,
	ChartDataSource,
	ChartShapeProperties,
	ChartSpaceSeries,
} from '../../chart/model-series';
import { drawingFillXml } from '../../drawingml/write-fill';
import type { DrawingFill } from '../../drawingml/types';
import { NS, parseXml } from '../../xml/index';
import type { ChartObject, ChartSeries } from '../model';
import { chartSeriesDrawingFill } from './chart-colors';
import { chartTitle } from './chart-title';

const [FIRST_AXIS, SECOND_AXIS] = DEFAULT_CHART_AXIS_IDS;

/** A string reference with its cache, or a literal: empty values keep their index, no point. */
function source(
	kind: 'str' | 'num',
	formula: string | undefined,
	values: readonly (string | number | null)[],
): ChartDataSource {
	const points = values.flatMap((value, index) =>
		value === null || value === '' ? [] : [{ index, value: String(value) }],
	);
	return {
		kind: formula ? (kind === 'str' ? 'strRef' : 'numRef') : kind === 'str' ? 'strLit' : 'numLit',
		...(formula ? { formula } : {}),
		cache: {
			type: kind === 'str' ? 'string' : 'number',
			...(kind === 'num' ? { formatCode: 'General' } : {}),
			pointCount: values.length,
			points,
		},
	};
}

function effects(series: ChartSeries): string | undefined {
	const xml = series.effectsXml;
	if (!xml) return undefined;
	const root = parseXml(xml).documentElement;
	if (root.namespaceURI !== NS.a || root.localName !== 'effectLst')
		throw new Error('Invalid chart effects XML');
	return xml;
}

/** Point overrides: a fill the shared writer can express, else the solid point colour. */
function dataPoints(series: ChartSeries): ChartDataPoint[] {
	const keys = new Set([
		...Object.keys(series.pointColors ?? {}),
		...Object.keys(series.pointFills ?? {}),
	]);
	return [...keys].flatMap((key) => {
		const fill = series.pointFills?.[Number(key)];
		const color = series.pointColors?.[Number(key)];
		const paint: DrawingFill | undefined = fill
			? drawingFillXml(fill)
				? fill
				: undefined
			: color
				? { kind: 'solid', color }
				: undefined;
		return /^\d+$/.test(key) && paint ? [{ index: Number(key), spPr: { fill: paint } }] : [];
	});
}

function seriesModel(chart: ChartObject, series: ChartSeries, index: number): ChartSpaceSeries {
	const type = chart.chartType;
	const out: ChartSpaceSeries = { index, order: index, dataPoints: dataPoints(series) };
	if (series.nameRef)
		out.tx = {
			reference: source('str', series.nameRef, series.name === undefined ? [] : [series.name]),
		};
	else if (series.name !== undefined) out.tx = { value: series.name };
	const lineLike = type === 'line' || type === 'scatter' || type === 'radar';
	const effectsXml = effects(series);
	if (
		(type !== 'pie' && type !== 'doughnut') ||
		series.fill ||
		series.color ||
		series.drawingColor ||
		effectsXml
	) {
		const fill = chartSeriesDrawingFill(chart, series, index);
		const shape: ChartShapeProperties = lineLike
			? { line: { widthEmu: 28575, cap: 'rnd', fill } }
			: { fill };
		if (effectsXml) shape.effectsXml = effectsXml;
		out.spPr = shape;
	}
	if (lineLike) out.marker = { symbol: 'none' };
	if (type === 'bar' || type === 'column') out.invertIfNegative = false;
	const numeric =
		series.categories.length > 0 && series.categories.every((c) => typeof c === 'number');
	const categories =
		series.categoriesRef !== undefined || series.categories.length > 0
			? source(numeric ? 'num' : 'str', series.categoriesRef, series.categories)
			: undefined;
	const values = source('num', series.valuesRef, series.values);
	if (type === 'scatter') {
		if (categories) out.xValues = categories;
		out.yValues = values;
		out.smooth = false;
		return out;
	}
	if (categories) out.categories = categories;
	out.values = values;
	if (type === 'line') out.smooth = false;
	return out;
}

/** The chart group and its axes for the chart type. */
function plot(chart: ChartObject, series: ChartSpaceSeries[]): [ChartPlotGroup, ChartAxis[]] {
	const grouping =
		chart.grouping ??
		(chart.chartType === 'bar' || chart.chartType === 'column' ? 'clustered' : 'standard');
	const flat = grouping === 'clustered' ? 'standard' : grouping;
	const axisIds = [FIRST_AXIS, SECOND_AXIS];
	const group = (kind: ChartPlotGroup['kind'], element: string, rest: Partial<ChartPlotGroup>) =>
		({ kind, element, is3D: false, series, axisIds, ...rest }) as ChartPlotGroup;
	switch (chart.chartType) {
		case 'line':
			return [
				group('line', 'lineChart', { grouping: flat, varyColors: false, marker: true }),
				defaultChartAxes(false, false),
			];
		case 'area':
			return [
				group('area', 'areaChart', { grouping: flat, varyColors: false }),
				defaultChartAxes(false, false),
			];
		case 'pie':
			return [group('pie', 'pieChart', { varyColors: true, firstSliceAngle: 0, axisIds: [] }), []];
		case 'doughnut':
			return [
				group('doughnut', 'doughnutChart', {
					varyColors: true,
					firstSliceAngle: 0,
					holeSize: 50,
					axisIds: [],
				}),
				[],
			];
		case 'scatter':
			return [
				group('scatter', 'scatterChart', { scatterStyle: 'lineMarker', varyColors: false }),
				defaultChartAxes(false, true),
			];
		case 'radar':
			return [
				group('radar', 'radarChart', { radarStyle: 'marker', varyColors: false }),
				defaultChartAxes(false, false),
			];
		default: {
			// Column and bar; the other families are generated as clustered columns.
			const horizontal = chart.chartType === 'bar';
			const bars = grouping === 'standard' ? 'clustered' : grouping;
			const overlap =
				chart.barOverlap === undefined && bars === 'clustered'
					? {}
					: { overlap: chart.barOverlap ?? 100 };
			return [
				group('bar', 'barChart', {
					barDirection: horizontal ? 'bar' : 'col',
					grouping: bars,
					varyColors: false,
					gapWidth: chart.barGapWidth ?? 150,
					...overlap,
				}),
				defaultChartAxes(horizontal, false),
			];
		}
	}
}

/** The built-in style: `c:style` 1..48, or the `c14:style` 101..148 choice with its fallback. */
function style(chart: ChartObject): Pick<ChartSpace, 'style' | 'c14Style'> {
	const id = effectiveBuiltInChartStyle(chart.formatting);
	if (!builtInChartStyleXml(id) || id === undefined) return {};
	return id > 100 ? { c14Style: id, style: id - 100 } : { style: id };
}

/** The neutral model of a new chart part for `chart`. */
export function chartSpaceFromObject(chart: ChartObject): ChartSpace {
	const series = chart.series.map((entry, index) => seriesModel(chart, entry, index));
	const [group, chartAxes] = plot(chart, series);
	const space: ChartSpace = {
		roundedCorners: false,
		...style(chart),
		autoTitleDeleted: chart.title === undefined,
		plotArea: { layout: {}, groups: [group], axes: chartAxes },
		plotVisibleOnly: true,
		displayBlanksAs: 'gap',
	};
	if (chart.title !== undefined) space.title = chartTitle(chart.title);
	if (chart.showLegend)
		space.legend = { position: chart.legendPosition ?? 'r', overlay: false, entries: [] };
	return withQuickLayout(space, chart.quickLayout);
}

/**
 * The parts of a session's Quick Layout the summary does not carry (axis titles, gridlines, value
 * axis, data labels, data table) applied to `space`; the title, legend and spacing come from the
 * summary fields. Unchanged when no layout was applied or the family takes none.
 */
export function withQuickLayout(space: ChartSpace, layout: number | undefined): ChartSpace {
	if (layout === undefined || !chartSupportsQuickLayout(space)) return space;
	return applyChartQuickLayout(space, layout, ['dataLabels', 'axes', 'dataTable']);
}
