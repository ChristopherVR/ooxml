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
} from '../../xml/index';
import type { ChartObject } from '../model';
import { parseChart } from '../read/chart';
import { chartXml, patchChartReferences } from './chart';
import { chartTitleXml } from './chart-title';
import { patchChartColors } from './chart-colors';
import { assertBarClusterOptions } from '../../chart/bar-cluster-geometry';
import { writeChartAxisFormatting } from '../../chart/write-axis-formatting';
import { writeChartFillFormatting } from '../../chart/write-fill-formatting';

type Doc = ReturnType<typeof parseXml>;

/** Parses a `c:` fragment and imports its root element into `doc`. */
function fragment(doc: Doc, xml: string): XmlElement {
	const wrapper = parseXml(`<w xmlns:c="${NS.c}" xmlns:a="${NS.a}">${xml}</w>`);
	const node = elements(wrapper.documentElement)[0];
	if (!node) throw new Error('Empty chart fragment');
	return doc.importNode(node, true) as XmlElement;
}

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
		const next = fragment(doc, chartTitleXml(title));
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
export function patchChartPart(
	xml: string,
	model: ChartObject,
	originalPalette?: number,
): string | undefined {
	assertBarClusterOptions(model);
	let before: ChartObject;
	try {
		before = parseChart(xml, model.anchor, model.partName ?? '');
		if (originalPalette !== undefined) before.colorPalette = originalPalette;
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
	const doc = parseXml(refs, { label: 'XLSX chart' });
	const formattingChanged = writeChartAxisFormatting(doc.documentElement, model.formatting);
	const fillsChanged = writeChartFillFormatting(doc.documentElement, model.formatting);
	const sameNames = before.series.every((s, i) => s.name === model.series[i]?.name);
	const sameSpacing =
		before.barGapWidth === model.barGapWidth && before.barOverlap === model.barOverlap;
	const sameColors = before.series.every(
		(s, i) =>
			JSON.stringify([
				s.color,
				s.drawingColor,
				s.pointColors,
				s.fill,
				s.pointFills,
				s.effectsXml,
			]) ===
			JSON.stringify([
				model.series[i]?.color,
				model.series[i]?.drawingColor,
				model.series[i]?.pointColors,
				model.series[i]?.fill,
				model.series[i]?.pointFills,
				model.series[i]?.effectsXml,
			]),
	);
	if (
		before.title === model.title &&
		before.showLegend === model.showLegend &&
		(model.legendPosition === undefined || before.legendPosition === model.legendPosition) &&
		sameNames &&
		sameSpacing &&
		sameColors &&
		!formattingChanged &&
		!fillsChanged &&
		before.colorPalette === model.colorPalette
	)
		return refs === xml ? undefined : refs;
	const chart = first(doc.documentElement, 'chart', NS.c);
	if (!chart) return refs === xml ? undefined : refs;
	if (before.title !== model.title) patchTitle(doc, chart, model.title);
	patchLegend(doc, chart, model);
	// Title/legend edits can create or replace their nodes after the initial change check.
	writeChartFillFormatting(doc.documentElement, model.formatting);
	if (!sameNames) patchSeriesNames(chart, model);
	if (!sameColors || before.colorPalette !== model.colorPalette)
		patchChartColors(doc, chart, model, before);
	if (!sameSpacing) {
		const plot =
			first(first(chart, 'plotArea', NS.c), 'barChart', NS.c) ??
			first(first(chart, 'plotArea', NS.c), 'bar3DChart', NS.c);
		if (plot)
			for (const [element, value, next] of [
				['gapWidth', model.barGapWidth, 'overlap'],
				['overlap', model.barOverlap, 'serLines'],
			] as const) {
				if (value === undefined) {
					const old = first(plot, element, NS.c);
					if (old) plot.removeChild(old);
				} else
					setVal(
						doc,
						plot,
						element,
						String(value),
						first(plot, next, NS.c) ?? first(plot, 'axId', NS.c),
					);
			}
	}
	return buildXml(doc);
}
