// Builds the painter's chart summary from an Office 2016 extended chart part (`cx:chartSpace`):
// the waterfall (`layoutId="waterfall"`) and the histogram (`layoutId="clusteredColumn"`). Word
// reads the cached points of `cx:data`; Excel files carry only the formula (a hidden
// `_xlchart.v1.N` defined name) and the host resolves it from live cells. Other extended families
// (funnel, treemap, box and whisker...) are reported by `layoutId` and not summarised.
import { parseDrawingFill } from '../../drawingml/drawing-fill';
import type { DrawingColor } from '../../drawingml/types';
import { NS, children, first, parseXml, type XmlElement } from '../../xml/index';
import type { ChartSummary, ChartSummarySeries, HistogramSummaryOptions } from './summary';

const cx = (parent: XmlElement | undefined, local: string): XmlElement | undefined =>
	first(parent, local, NS.cx);
const text = (element: XmlElement | undefined): string => (element?.textContent ?? '').trim();

/** `@val` when written, else the element's text (the schema uses both for scalar children). */
function scalar(parent: XmlElement | undefined, local: string): string | undefined {
	const element = cx(parent, local);
	if (!element) return undefined;
	return element.hasAttribute('val') ? (element.getAttribute('val') ?? '') : text(element);
}

/** The data of a series: its own `cx:data`, else the shared one named by `cx:dataId`. */
function seriesData(root: XmlElement, series: XmlElement): XmlElement | undefined {
	const inline = cx(series, 'data');
	if (inline) return inline;
	const id = cx(series, 'dataId')?.getAttribute('val');
	const shared = children(cx(root, 'chartData') ?? root, 'data', NS.cx);
	return id === null || id === undefined
		? shared[0]
		: shared.find((d) => d.getAttribute('id') === id);
}

interface Dimension {
	formula?: string;
	points: string[];
}

/** One dimension's formula and its first level of cached points, holes as `''`. */
function dimension(data: XmlElement | undefined, kind: 'strDim' | 'numDim', type: string) {
	const all = data ? children(data, kind, NS.cx) : [];
	const element = all.find((d) => (d.getAttribute('type') ?? '') === type) ?? undefined;
	if (!element) return undefined;
	const out: Dimension = { points: [] };
	const formula = text(cx(element, 'f'));
	if (formula) out.formula = formula;
	const level = cx(element, 'lvl');
	const count = Number(level?.getAttribute('ptCount') ?? 0);
	for (const point of level ? children(level, 'pt', NS.cx) : []) {
		const index = Number(point.getAttribute('idx') ?? out.points.length);
		if (Number.isInteger(index) && index >= 0) out.points[index] = text(point);
	}
	for (let i = 0; i < Math.max(count, out.points.length); i++) out.points[i] ??= '';
	return out;
}

/** The title text: `cx:v` of `cx:txData`, or the runs of `cx:rich`. */
function titleOf(chart: XmlElement | undefined): string | undefined {
	const tx = cx(cx(chart, 'title'), 'tx');
	const data = text(cx(cx(tx, 'txData'), 'v'));
	if (data) return data;
	const rich = cx(tx, 'rich');
	if (!rich) return undefined;
	const runs = Array.from(rich.getElementsByTagNameNS(NS.a, 't')).map((t) => t.textContent ?? '');
	return runs.join('') || undefined;
}

function seriesFill(series: XmlElement): Pick<ChartSummarySeries, 'drawingColor'> {
	const fill = parseDrawingFill(cx(series, 'spPr'));
	const color: DrawingColor | undefined = fill?.kind === 'solid' ? fill.color : undefined;
	return color ? { drawingColor: color } : {};
}

function numberOf(value: string | undefined): number | undefined {
	const parsed = Number.parseFloat(value ?? '');
	return Number.isFinite(parsed) ? parsed : undefined;
}

function histogramOptions(layoutPr: XmlElement | undefined): HistogramSummaryOptions {
	const binning = cx(layoutPr, 'binning');
	const out: HistogramSummaryOptions = {};
	const size = numberOf(scalar(binning, 'binSize'));
	const count = numberOf(scalar(binning, 'binCount'));
	if (size !== undefined) out.binSize = size;
	if (count !== undefined && count >= 0) out.binCount = Math.trunc(count);
	const closed = binning?.getAttribute('intervalClosed');
	if (closed === 'l' || closed === 'r') out.intervalClosed = closed;
	for (const key of ['underflow', 'overflow'] as const) {
		const raw = binning?.getAttribute(key);
		if (raw === 'auto') out[key] = 'auto';
		else if (raw !== null && raw !== undefined && numberOf(raw) !== undefined)
			out[key] = numberOf(raw);
	}
	if (cx(layoutPr, 'aggregation')) out.aggregateByCategory = true;
	return out;
}

/** What an extended chart part is: its layout and, for the families the painter draws, a summary. */
export interface ChartExSummary {
	/** `cx:series/@layoutId` of the first series (`waterfall`, `clusteredColumn`, `funnel`...). */
	layoutId: string;
	summary?: ChartSummary;
}

const LEGEND_POSITIONS = ['r', 'l', 't', 'b', 'tr'];

/** The summary of a `cx:chartSpace`, or undefined when the part is not an extended chart. */
export function chartExSummary(source: string | XmlElement): ChartExSummary | undefined {
	const root =
		typeof source === 'string' ? parseXml(source, { label: 'Chart part' }).documentElement : source;
	if (root.localName !== 'chartSpace' || root.namespaceURI !== NS.cx) return undefined;
	const chart = cx(root, 'chart');
	const region = cx(cx(chart, 'plotArea'), 'plotAreaRegion');
	const all = region ? children(region, 'series', NS.cx) : [];
	const primary = all.find((s) => s.getAttribute('layoutId') !== 'paretoLine') ?? all[0];
	const layoutId = primary?.getAttribute('layoutId') ?? '';
	const chartType =
		layoutId === 'waterfall'
			? 'waterfall'
			: layoutId === 'clusteredColumn'
				? 'histogram'
				: undefined;
	if (!primary || !chartType) return { layoutId };
	const data = seriesData(root, primary);
	const categories = dimension(data, 'strDim', 'cat');
	const values = dimension(data, 'numDim', 'val');
	const txData = cx(cx(primary, 'tx'), 'txData');
	const series: ChartSummarySeries = {
		categories: categories?.points ?? [],
		values: (values?.points ?? []).map((p) =>
			p !== '' && Number.isFinite(Number(p)) ? Number(p) : null,
		),
		...seriesFill(primary),
	};
	const name = text(cx(txData, 'v'));
	if (name) series.name = name;
	const nameRef = text(cx(txData, 'f'));
	if (nameRef) series.nameRef = nameRef;
	if (categories?.formula) series.categoriesRef = categories.formula;
	if (values?.formula) series.valuesRef = values.formula;
	const layoutPr = cx(primary, 'layoutPr');
	const legend = cx(chart, 'legend');
	const summary: ChartSummary = { chartType, series: [series], showLegend: legend !== undefined };
	const title = titleOf(chart);
	if (title) summary.title = title;
	const position = legend?.getAttribute('pos');
	if (position && LEGEND_POSITIONS.includes(position))
		summary.legendPosition = position as NonNullable<ChartSummary['legendPosition']>;
	if (chartType === 'waterfall') {
		const subtotals = cx(layoutPr, 'subtotals');
		const connectors = cx(layoutPr, 'visibility')?.getAttribute('connectorLines');
		summary.waterfall = {
			...(subtotals
				? {
						subtotalIndices: children(subtotals, 'idx', NS.cx)
							.map((idx) => Number.parseInt(idx.getAttribute('val') ?? '', 10))
							.filter((idx) => Number.isInteger(idx) && idx >= 0),
					}
				: {}),
			...(connectors === '0' || connectors === 'false' ? { connectorLines: false } : {}),
		};
	} else summary.histogram = histogramOptions(layoutPr);
	return { layoutId, summary };
}
