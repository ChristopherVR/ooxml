// `c:chartSpace` parser: a series' trendlines (`c:trendline`, with `c:trendlineLbl`) and error bars
// (`c:errBars`). Every element is kept, even one whose type cannot be read, so a product that pairs
// them with its own tree by position stays aligned.
import { NS, children, type XmlElement } from '../xml/index';
import {
	CHART_ERROR_BAR_DIRECTIONS,
	CHART_ERROR_BAR_TYPES,
	CHART_ERROR_VALUE_TYPES,
	type ChartErrorBars,
	type ChartTrendline,
	type ChartTrendlineLabel,
	type ChartTrendlineType,
} from './model-analysis';
import { parseChartText, parseDataSource } from './parse-text';
import {
	assignDefined,
	cBool,
	cChild,
	cEnum,
	cNumber,
	extensionList,
	manualLayout,
	numberFormat,
	reportUnmodelled,
	shapeProperties,
	textProperties,
	type ChartParseContext,
} from './parse-util';

const TRENDLINE_TYPES: readonly ChartTrendlineType[] = [
	'exp',
	'linear',
	'log',
	'movingAvg',
	'poly',
	'power',
];

const LABEL = new Set(['layout', 'tx', 'numFmt', 'spPr', 'txPr', 'extLst']);

function parseTrendlineLabel(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartTrendlineLabel | undefined {
	if (!element) return undefined;
	reportUnmodelled(context, element, LABEL);
	return assignDefined<ChartTrendlineLabel>(
		{},
		{
			layout: manualLayout(element),
			tx: parseChartText(context, cChild(element, 'tx'), false),
			numberFormat: numberFormat(element),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			extLst: extensionList(element),
		},
	);
}

const TRENDLINE = new Set([
	'name',
	'spPr',
	'trendlineType',
	'order',
	'period',
	'forward',
	'backward',
	'intercept',
	'dispRSqr',
	'dispEq',
	'trendlineLbl',
	'extLst',
]);

/** One trendline (`c:trendline`). */
export function parseTrendline(context: ChartParseContext, element: XmlElement): ChartTrendline {
	reportUnmodelled(context, element, TRENDLINE);
	const name = cChild(element, 'name');
	return assignDefined<ChartTrendline>(
		{},
		{
			name: name ? (name.textContent ?? '') : undefined,
			spPr: shapeProperties(element),
			trendlineType: cEnum(context, element, 'trendlineType', TRENDLINE_TYPES),
			order: cNumber(context, element, 'order'),
			period: cNumber(context, element, 'period'),
			forward: cNumber(context, element, 'forward'),
			backward: cNumber(context, element, 'backward'),
			intercept: cNumber(context, element, 'intercept'),
			displayRSquared: cBool(context, element, 'dispRSqr'),
			displayEquation: cBool(context, element, 'dispEq'),
			label: parseTrendlineLabel(context, cChild(element, 'trendlineLbl')),
			extLst: extensionList(element),
		},
	);
}

const ERROR_BARS = new Set([
	'errDir',
	'errBarType',
	'errValType',
	'noEndCap',
	'plus',
	'minus',
	'val',
	'spPr',
	'extLst',
]);

/** One set of error bars (`c:errBars`). */
export function parseErrorBars(context: ChartParseContext, element: XmlElement): ChartErrorBars {
	reportUnmodelled(context, element, ERROR_BARS);
	return assignDefined<ChartErrorBars>(
		{},
		{
			direction: cEnum(context, element, 'errDir', CHART_ERROR_BAR_DIRECTIONS),
			barType: cEnum(context, element, 'errBarType', CHART_ERROR_BAR_TYPES),
			valueType: cEnum(context, element, 'errValType', CHART_ERROR_VALUE_TYPES),
			noEndCap: cBool(context, element, 'noEndCap'),
			plus: parseDataSource(cChild(element, 'plus')),
			minus: parseDataSource(cChild(element, 'minus')),
			value: cNumber(context, element, 'val'),
			spPr: shapeProperties(element),
			extLst: extensionList(element),
		},
	);
}

/** The trendlines and error bars of a series, each absent when none are written. */
export function parseSeriesAnalysis(
	context: ChartParseContext,
	series: XmlElement,
): { trendlines?: ChartTrendline[]; errorBars?: ChartErrorBars[] } {
	const trendlines = children(series, 'trendline', NS.c).map((node) =>
		parseTrendline(context, node),
	);
	const errorBars = children(series, 'errBars', NS.c).map((node) => parseErrorBars(context, node));
	return {
		...(trendlines.length ? { trendlines } : {}),
		...(errorBars.length ? { errorBars } : {}),
	};
}
