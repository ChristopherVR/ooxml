/**
 * The trendlines and error bars of a pptx chart series from the neutral model
 * (`ChartSpaceSeries.trendlines`, `.errorBars`, parsed by `chart/parse-analysis.ts`). The
 * structural fields (type, name, order, period, extensions, intercept, display switches, label
 * number format, direction, bar and value type, end cap, value and custom amounts) come from the
 * neutral model. Line paint, width and dash use neutral DrawingML fields and the shared color
 * resolver. Only opaque manual-layout extensions still come from the matching source node.
 *
 * @module chart-from-neutral-analysis
 */
import {
	chartTrendlineKind,
	type ChartDataSource,
	type ChartErrorBars,
	type ChartTrendline,
} from '../../../chart/index';
import type {
	PptxChartErrBars,
	PptxChartTrendline,
	PptxChartTrendlineLabel,
	XmlObject,
} from '../types';
import { chartLineFromNeutral } from './chart-from-neutral-paint';
import { parseChartManualLayout } from './chart-layout';
import type { ChartTreeReader } from './chart-series-tree-fields';

function labelFromNeutral(
	label: ChartTrendline['label'],
	node: XmlObject | undefined,
	reader: ChartTreeReader,
): PptxChartTrendlineLabel | undefined {
	if (!label) return undefined;
	const result: PptxChartTrendlineLabel = {};
	const { sourceXml: _source, ...values } = label.layout ?? {};
	const extension = parseChartManualLayout(
		reader.xml.getChildByLocalName(node, 'trendlineLbl'),
		reader.localName,
	)?.ext;
	const layout = { ...values, ...(extension ? { ext: extension } : {}) };
	if (layout && Object.keys(layout).length) result.layout = { ...layout };
	if (label.numberFormat) result.numberFormatCode = label.numberFormat.formatCode;
	if (label.numberFormat?.sourceLinked !== undefined)
		result.sourceLinked = label.numberFormat.sourceLinked;
	return result;
}

/** One pptx trendline; `undefined` when its type cannot be read (pptx drops it). */
function trendlineFromNeutral(
	trendline: ChartTrendline,
	node: XmlObject | undefined,
	reader: ChartTreeReader,
): PptxChartTrendline | undefined {
	const trendlineType = chartTrendlineKind(trendline.trendlineType);
	if (!trendlineType) return undefined;
	const result: PptxChartTrendline = { trendlineType };
	if (trendline.name !== undefined) result.name = trendline.name;
	const { order, period } = trendline;
	if (order !== undefined && Number.isInteger(order) && order >= 2 && order <= 6)
		result.order = order;
	if (period !== undefined && Number.isInteger(period) && period >= 2) result.period = period;
	if (trendline.forward !== undefined) result.forward = trendline.forward;
	if (trendline.backward !== undefined) result.backward = trendline.backward;
	if (trendline.intercept !== undefined) result.intercept = trendline.intercept;
	if (trendline.displayRSquared !== undefined) result.displayRSq = trendline.displayRSquared;
	if (trendline.displayEquation !== undefined) result.displayEq = trendline.displayEquation;
	const line = chartLineFromNeutral(trendline.spPr?.line, reader);
	if (line.color) result.color = line.color;
	if (line.width !== undefined) result.lineWidth = line.width;
	if (line.dashStyle) result.lineDashStyle = line.dashStyle;
	const label = labelFromNeutral(trendline.label, node, reader);
	if (label) result.label = label;
	return result;
}

/** The trendlines of a series; `seriesNode` is its `c:ser` in the object tree. */
export function trendlinesFromNeutral(
	trendlines: readonly ChartTrendline[] | undefined,
	seriesNode: XmlObject,
	reader: ChartTreeReader,
): PptxChartTrendline[] {
	if (!trendlines?.length) return [];
	const nodes = reader.xml.getChildrenArrayByLocalName(seriesNode, 'trendline');
	return trendlines
		.map((trendline, index) => trendlineFromNeutral(trendline, nodes[index], reader))
		.filter((trendline): trendline is PptxChartTrendline => trendline !== undefined);
}

/**
 * The custom amounts of `c:plus` or `c:minus`: the numeric cache by point index, blanks and
 * unreadable points dropped (dense, as pptx has always held them).
 */
function customAmounts(source: ChartDataSource | undefined): number[] {
	if (source?.kind !== 'numRef' && source?.kind !== 'numLit') return [];
	return [...(source.cache?.points ?? [])]
		.sort((left, right) => (left.index ?? 0) - (right.index ?? 0))
		.map((point) => point.value.trim())
		.filter((value) => value.length > 0)
		.map((value) => parseFloat(value))
		.filter((value) => Number.isFinite(value));
}

function errBarsFromNeutralOne(
	bars: ChartErrorBars,
	node: XmlObject | undefined,
	reader: ChartTreeReader,
): PptxChartErrBars | undefined {
	if (!bars.valueType) return undefined;
	const result: PptxChartErrBars = {
		direction: bars.direction === 'x' ? 'x' : 'y',
		barType: bars.barType ?? 'both',
		valType: bars.valueType,
	};
	if (bars.noEndCap !== undefined) result.noEndCap = bars.noEndCap;
	if (bars.value !== undefined) result.val = bars.value;
	if (bars.valueType === 'cust') {
		const plus = customAmounts(bars.plus);
		if (plus.length > 0) result.customPlus = plus;
		const minus = customAmounts(bars.minus);
		if (minus.length > 0) result.customMinus = minus;
	}
	const line = chartLineFromNeutral(bars.spPr?.line, reader);
	if (line.color) result.color = line.color;
	if (line.width !== undefined) result.width = line.width;
	if (line.dashStyle) result.dashStyle = line.dashStyle;
	return result;
}

/** The error bars of a series; `seriesNode` is its `c:ser` in the object tree. */
export function errBarsFromNeutral(
	errorBars: readonly ChartErrorBars[] | undefined,
	seriesNode: XmlObject,
	reader: ChartTreeReader,
): PptxChartErrBars[] {
	if (!errorBars?.length) return [];
	const nodes = reader.xml.getChildrenArrayByLocalName(seriesNode, 'errBars');
	return errorBars
		.map((bars, index) => errBarsFromNeutralOne(bars, nodes[index], reader))
		.filter((bars): bars is PptxChartErrBars => bars !== undefined);
}
