import { NS, children, elements, first, parseXml, type XmlElement } from '../../xml/index';
import type { ChartObject, ChartSeries, ChartType, Color, DrawingAnchor } from '../model';
import { att } from './xml-util';
import { parseDrawingColorIn } from '../../diagram/drawing-color';
import { readChartFormatting } from '../../chart/read-formatting';
import { parseDrawingFill } from '../../diagram/drawing-fill';
import { buildXml } from '../../xml/index';
import { parseDrawingTextBody } from '../../diagram/drawing-text';

const c = (parent: ParentNode | null | undefined, local: string) => first(parent, local, NS.c);
const val = (parent: ParentNode | null | undefined, local: string) => att(c(parent, local), 'val');

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

/** The solid fill colour of a `c:spPr`, if it has a plain one. */
export function solidFillColor(spPr: XmlElement | undefined): Color | undefined {
	const fill = first(spPr, 'solidFill', NS.a);
	const node = fill ? elements(fill)[0] : undefined;
	if (!node) return undefined;
	if (node.localName === 'srgbClr') {
		const rgb = att(node, 'val');
		return rgb ? { rgb: rgb.toUpperCase() } : undefined;
	}
	if (node.localName === 'schemeClr') {
		const index = SCHEME_INDEX[att(node, 'val') ?? ''];
		return index === undefined ? undefined : { theme: index };
	}
	return undefined;
}

/** Plain text of a rich text body (`c:rich` / `c:tx`), paragraphs joined with newlines. */
function richText(tx: XmlElement | undefined): string | undefined {
	const rich = c(tx, 'rich');
	if (rich) {
		return parseDrawingTextBody(rich)?.text;
	}
	const cached = cache(c(c(tx, 'strRef'), 'strCache'));
	return cached[0] === undefined ? undefined : String(cached[0]);
}

function cache(node: XmlElement | undefined): (string | number)[] {
	if (!node) return [];
	const count = Number(val(node, 'ptCount') ?? 0);
	const out: (string | number)[] = [];
	for (const pt of children(node, 'pt', NS.c)) {
		const index = Number(att(pt, 'idx') ?? out.length);
		const text = c(pt, 'v')?.textContent ?? '';
		const numeric = node.localName === 'numCache' || node.localName === 'numLit';
		out[index] =
			numeric && text.trim() !== '' && Number.isFinite(Number(text)) ? Number(text) : text;
	}
	const size = Math.max(count, out.length);
	for (let i = 0; i < size; i++) if (out[i] === undefined) out[i] = '';
	return out;
}

/** A data source (`c:cat`, `c:val`, `c:xVal`...): its formula and cached points. */
function source(node: XmlElement | undefined): { ref?: string; points: (string | number)[] } {
	const ref = c(node, 'numRef') ?? c(node, 'strRef') ?? c(node, 'multiLvlStrRef');
	const f = c(ref, 'f')?.textContent ?? undefined;
	const points = ref
		? cache(c(ref, 'numCache') ?? c(ref, 'strCache') ?? c(c(ref, 'multiLvlStrCache'), 'lvl'))
		: cache(c(node, 'numLit') ?? c(node, 'strLit'));
	const out: { ref?: string; points: (string | number)[] } = { points };
	if (f) out.ref = f;
	return out;
}

const PLOT_TYPES: Record<string, ChartType> = {
	barChart: 'bar',
	bar3DChart: 'bar',
	lineChart: 'line',
	line3DChart: 'line',
	pieChart: 'pie',
	pie3DChart: 'pie',
	ofPieChart: 'pie',
	doughnutChart: 'doughnut',
	areaChart: 'area',
	area3DChart: 'area',
	scatterChart: 'scatter',
	radarChart: 'radar',
	bubbleChart: 'bubble',
	stockChart: 'stock',
	surfaceChart: 'surface',
	surface3DChart: 'surface',
};

const GROUPINGS = new Set(['clustered', 'stacked', 'percentStacked', 'standard']);

/** Reads the modelled summary of a `c:chartSpace` part. */
export function parseChart(
	xml: string,
	anchor: DrawingAnchor,
	partName: string,
	name?: string,
): ChartObject {
	const root = parseXml(xml, { label: 'XLSX chart' }).documentElement;
	const chart = c(root, 'chart');
	const plotArea = c(chart, 'plotArea');
	const plot = plotArea
		? elements(plotArea).find((node) => PLOT_TYPES[node.localName ?? ''])
		: undefined;
	let chartType: ChartType = plot ? (PLOT_TYPES[plot.localName ?? ''] ?? 'column') : 'column';
	if (chartType === 'bar' && val(plot, 'barDir') !== 'bar') chartType = 'column';
	const series: ChartSeries[] = (plot ? children(plot, 'ser', NS.c) : []).map((ser) => {
		const tx = c(ser, 'tx');
		const nameRef = c(c(tx, 'strRef'), 'f')?.textContent ?? undefined;
		const categories = source(c(ser, 'cat') ?? c(ser, 'xVal'));
		const values = source(c(ser, 'val') ?? c(ser, 'yVal'));
		const out: ChartSeries = {
			categories: categories.points,
			values: values.points.map((v) =>
				typeof v === 'number' ? v : v === '' ? null : Number.isFinite(Number(v)) ? Number(v) : null,
			),
		};
		const seriesName = c(tx, 'v')?.textContent ?? richText(tx);
		if (seriesName !== undefined) out.name = seriesName;
		if (nameRef) out.nameRef = nameRef;
		if (categories.ref) out.categoriesRef = categories.ref;
		if (values.ref) out.valuesRef = values.ref;
		const spPr = c(ser, 'spPr');
		const effects = first(spPr, 'effectLst', NS.a);
		if (effects) out.effectsXml = buildXml(effects);
		const primary = ['line', 'scatter', 'radar'].includes(chartType)
			? first(spPr, 'ln', NS.a)
			: spPr;
		const colorContainer =
			chartType === 'scatter' && !first(primary, 'solidFill', NS.a)
				? c(c(ser, 'marker'), 'spPr')
				: primary;
		const color = solidFillColor(colorContainer);
		if (color) out.color = color;
		const drawingColor = parseDrawingColorIn(first(colorContainer, 'solidFill', NS.a));
		if (drawingColor) out.drawingColor = drawingColor;
		const fill = parseDrawingFill(colorContainer);
		if (fill && fill.kind !== 'solid' && !(chartType === 'scatter' && fill.kind === 'none'))
			out.fill = fill;
		const points: NonNullable<ChartSeries['pointColors']> = {};
		const pointFills: NonNullable<ChartSeries['pointFills']> = {};
		for (const point of children(ser, 'dPt', NS.c)) {
			const pointColor = parseDrawingColorIn(
				first(c(point, 'spPr') ?? c(c(point, 'marker'), 'spPr'), 'solidFill', NS.a),
			);
			const index = Number(val(point, 'idx'));
			if (pointColor && Number.isInteger(index) && index >= 0) points[index] = pointColor;
			const pointFill = parseDrawingFill(c(point, 'spPr') ?? c(c(point, 'marker'), 'spPr'));
			if (pointFill && pointFill.kind !== 'solid' && Number.isInteger(index) && index >= 0)
				pointFills[index] = pointFill;
		}
		if (Object.keys(points).length) out.pointColors = points;
		if (Object.keys(pointFills).length) out.pointFills = pointFills;
		return out;
	});
	const object: ChartObject = {
		kind: 'chart',
		anchor,
		chartType,
		series,
		showLegend: c(chart, 'legend') !== undefined,
		partName,
	};
	const grouping = val(plot, 'grouping');
	if (grouping && GROUPINGS.has(grouping))
		object.grouping = grouping as NonNullable<ChartObject['grouping']>;
	if (chartType === 'bar' || chartType === 'column') {
		for (const [element, field, min, max] of [
			['gapWidth', 'barGapWidth', 0, 500],
			['overlap', 'barOverlap', -100, 100],
		] as const) {
			const raw = val(plot, element);
			const value =
				raw === undefined ? (field === 'barGapWidth' ? 150 : 0) : !raw.trim() ? NaN : Number(raw);
			if (Number.isInteger(value) && value >= min && value <= max) object[field] = value;
		}
	}
	const titleNode = c(chart, 'title');
	if (titleNode && val(chart, 'autoTitleDeleted') !== '1') {
		const title = richText(c(titleNode, 'tx'));
		if (title !== undefined) object.title = title;
		else if (series.length === 1 && series[0]?.name) object.title = series[0].name;
	}
	const legendPos = val(c(chart, 'legend'), 'legendPos');
	if (legendPos && ['r', 'l', 't', 'b', 'tr'].includes(legendPos))
		object.legendPosition = legendPos as NonNullable<ChartObject['legendPosition']>;
	if (name) object.name = name;
	const formatting = readChartFormatting(root);
	if (formatting) object.formatting = formatting;
	return object;
}
