// Patching a kept chart part to the model: title, legend, series names and references are
// edited in place so unmodelled formatting survives; a changed type or series list regenerates.
import {
	NS,
	buildXml,
	children,
	elements,
	first,
	parseXml,
	type XmlElement,
} from '../../xml/index.js';
import type { ChartObject } from '../model.js';
import { parseChart } from '../read/chart.js';
import { chartXml, patchChartReferences } from './chart.js';
import { escapeText } from './xml-out.js';

type Doc = ReturnType<typeof parseXml>;

/** Parses a `c:` fragment and imports its root element into `doc`. */
function fragment(doc: Doc, xml: string): XmlElement {
	const wrapper = parseXml(`<w xmlns:c="${NS.c}" xmlns:a="${NS.a}">${xml}</w>`);
	const node = elements(wrapper.documentElement)[0];
	if (!node) throw new Error('Empty chart fragment');
	return doc.importNode(node, true) as XmlElement;
}

const titleXml = (title: string) =>
	`<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeText(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;

function setVal(doc: Doc, parent: XmlElement, local: string, value: string, before?: XmlElement) {
	let node = first(parent, local, NS.c);
	if (!node) {
		node = fragment(doc, `<c:${local} val="${value}"/>`);
		parent.insertBefore(node, before ?? null);
	}
	node.setAttribute('val', value);
}

function patchTitle(doc: Doc, chart: XmlElement, title: string | undefined): void {
	const existing = first(chart, 'title', NS.c);
	if (title === undefined) {
		if (existing) chart.removeChild(existing);
		setVal(doc, chart, 'autoTitleDeleted', '1', first(chart, 'plotArea', NS.c));
		return;
	}
	const runs = existing
		? [...existing.getElementsByTagNameNS(NS.a, 'r')].filter((r) => r.parentNode)
		: [];
	const firstText = runs[0]?.getElementsByTagNameNS(NS.a, 't')[0];
	if (existing && firstText && !first(first(existing, 'tx', NS.c), 'strRef', NS.c)) {
		firstText.textContent = title;
		for (const run of runs.slice(1)) run.parentNode?.removeChild(run);
	} else {
		const next = fragment(doc, titleXml(title));
		if (existing) chart.replaceChild(next, existing);
		else chart.insertBefore(next, elements(chart)[0] ?? null);
	}
	const head = first(chart, 'title', NS.c);
	setVal(doc, chart, 'autoTitleDeleted', '0', head?.nextSibling as XmlElement | undefined);
}

function patchLegend(doc: Doc, chart: XmlElement, model: ChartObject): void {
	const legend = first(chart, 'legend', NS.c);
	if (!model.showLegend) {
		if (legend) chart.removeChild(legend);
		return;
	}
	const pos = model.legendPosition ?? 'r';
	if (!legend) {
		const node = fragment(
			doc,
			`<c:legend><c:legendPos val="${pos}"/><c:overlay val="0"/></c:legend>`,
		);
		const plotArea = first(chart, 'plotArea', NS.c);
		chart.insertBefore(node, plotArea?.nextSibling ?? null);
		return;
	}
	if (model.legendPosition) setVal(doc, legend, 'legendPos', pos, elements(legend)[0]);
}

function patchSeriesNames(chart: XmlElement, model: ChartObject): void {
	const plotArea = first(chart, 'plotArea', NS.c);
	const plot = plotArea
		? elements(plotArea).find((node) => children(node, 'ser', NS.c).length)
		: undefined;
	if (!plot) return;
	children(plot, 'ser', NS.c).forEach((ser, i) => {
		const name = model.series[i]?.name;
		const v = first(first(ser, 'tx', NS.c), 'v', NS.c);
		if (v && name !== undefined && !model.series[i]?.nameRef) v.textContent = name;
	});
}

/**
 * The chart part to write for a kept chart, or `undefined` when the source part can be copied
 * unchanged. Edits the source XML in place when only the title, legend, series names or series
 * references changed; returns a regenerated part (losing unmodelled detail) when the chart type,
 * grouping or number of series changed.
 */
export function patchChartPart(xml: string, model: ChartObject): string | undefined {
	let before: ChartObject;
	try {
		before = parseChart(xml, model.anchor, model.partName ?? '');
	} catch {
		return undefined;
	}
	if (
		before.chartType !== model.chartType ||
		(model.grouping !== undefined && before.grouping !== model.grouping) ||
		before.series.length !== model.series.length
	)
		return chartXml(model);
	const refs = patchChartReferences(xml, model) ?? xml;
	const sameNames = before.series.every((s, i) => s.name === model.series[i]?.name);
	if (
		before.title === model.title &&
		before.showLegend === model.showLegend &&
		(model.legendPosition === undefined || before.legendPosition === model.legendPosition) &&
		sameNames
	)
		return refs === xml ? undefined : refs;
	const doc = parseXml(refs, { label: 'XLSX chart' });
	const chart = first(doc.documentElement, 'chart', NS.c);
	if (!chart) return refs === xml ? undefined : refs;
	if (before.title !== model.title) patchTitle(doc, chart, model.title);
	patchLegend(doc, chart, model);
	if (!sameNames) patchSeriesNames(chart, model);
	return buildXml(doc);
}
