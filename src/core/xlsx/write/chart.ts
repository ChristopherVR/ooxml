import { NS, buildXml, children, elements, first, parseXml } from '../../xml/index';
import type { ChartObject } from '../model';
import { assertBarClusterOptions } from '../../chart/bar-cluster-geometry';
import { writeChartAxisFormatting } from '../../chart/write-axis-formatting';
import { writeChartFillFormatting } from '../../chart/write-fill-formatting';
import { writeChartLayoutFormatting } from '../../chart/write-layout-formatting';
import { writeChartSpace } from '../../chart/write-space';
import { writeChartTextFormatting } from '../../chart/write-text-formatting';
import { chartSpaceFromObject } from './chart-space';

/**
 * A new chart part for a chart created in the model (no source part to keep): the neutral model
 * of the chart (`chart-space.ts`) through the shared writer, then the direct formatting patches.
 */
export function chartXml(chart: ChartObject): string {
	assertBarClusterOptions(chart);
	const xml = writeChartSpace(chartSpaceFromObject(chart), { declarationBreak: '\n' });
	if (!chart.formatting) return xml;
	const doc = parseXml(xml);
	const textChanged = writeChartTextFormatting(doc.documentElement, chart.formatting);
	const axesChanged = writeChartAxisFormatting(doc.documentElement, chart.formatting);
	const fillsChanged = writeChartFillFormatting(doc.documentElement, chart.formatting);
	const layoutChanged = writeChartLayoutFormatting(doc.documentElement, chart.formatting);
	return textChanged || axesChanged || fillsChanged || layoutChanged ? buildXml(doc) : xml;
}

/**
 * Updates the series references of a kept chart part when the model moved them (rows inserted,
 * sheet renamed). Returns `undefined` when nothing changed, so the part stays byte-identical.
 */
export function patchChartReferences(xml: string, chart: ChartObject): string | undefined {
	const doc = parseXml(xml, { label: 'XLSX chart' });
	const plotArea = first(first(doc.documentElement, 'chart', NS.c), 'plotArea', NS.c);
	const plot = plotArea
		? elements(plotArea).find((node) => children(node, 'ser', NS.c).length)
		: undefined;
	if (!plot) return undefined;
	let changed = false;
	const setRef = (node: Element | undefined, ref: string | undefined) => {
		const f = node
			? first(first(node, 'numRef', NS.c) ?? first(node, 'strRef', NS.c), 'f', NS.c)
			: undefined;
		if (f && ref !== undefined && f.textContent !== ref) {
			f.textContent = ref;
			changed = true;
		}
	};
	children(plot, 'ser', NS.c).forEach((ser, index) => {
		const model = chart.series[index];
		if (!model) return;
		setRef(first(ser, 'tx', NS.c), model.nameRef);
		setRef(first(ser, 'cat', NS.c) ?? first(ser, 'xVal', NS.c), model.categoriesRef);
		setRef(first(ser, 'val', NS.c) ?? first(ser, 'yVal', NS.c), model.valuesRef);
	});
	return changed ? buildXml(doc) : undefined;
}
