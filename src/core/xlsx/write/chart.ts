import { buildXml, parseXml } from '../../xml/index';
import type { ChartSpace } from '../../chart/model';
import { parseChartSpace } from '../../chart/parse-space';
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
 * The neutral model of the part {@link chartXml} writes for `chart`: what a chart built from the
 * summary alone (a new chart, or a type the operations cannot switch to) is edited and saved as.
 */
export function newChartSpaceOf(chart: ChartObject): ChartSpace {
	return parseChartSpace(chartXml(chart)).chartSpace;
}
