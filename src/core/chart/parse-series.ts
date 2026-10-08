// `c:chartSpace` parser: series (`c:ser`), their markers and data point overrides.
import { NS, children, type XmlElement } from '../xml/index';
import type { ChartDataPoint, ChartMarker, ChartSpaceSeries } from './model-series';
import { parseChartText, parseDataLabels, parseDataSource } from './parse-text';
import {
	assignDefined,
	cBool,
	cChild,
	cNumber,
	cVal,
	extensionList,
	reportUnmodelled,
	shapeProperties,
	type ChartParseContext,
} from './parse-util';

const MARKER = new Set(['symbol', 'size', 'spPr', 'extLst']);

/** A marker (`c:marker` of a series or data point). */
export function parseMarker(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartMarker | undefined {
	if (!element) return undefined;
	reportUnmodelled(context, element, MARKER);
	return assignDefined<ChartMarker>(
		{},
		{
			symbol: cVal(element, 'symbol'),
			size: cNumber(context, element, 'size'),
			spPr: shapeProperties(element),
			extLst: extensionList(element),
		},
	);
}

const POINT = new Set([
	'idx',
	'invertIfNegative',
	'marker',
	'bubble3D',
	'explosion',
	'spPr',
	'extLst',
]);

function parseDataPoint(context: ChartParseContext, element: XmlElement): ChartDataPoint {
	reportUnmodelled(context, element, POINT);
	const index = cNumber(context, element, 'idx');
	return assignDefined<ChartDataPoint>(
		{},
		{
			index: index !== undefined && Number.isInteger(index) && index >= 0 ? index : undefined,
			invertIfNegative: cBool(context, element, 'invertIfNegative'),
			bubble3D: cBool(context, element, 'bubble3D'),
			explosion: cNumber(context, element, 'explosion'),
			marker: parseMarker(context, cChild(element, 'marker')),
			spPr: shapeProperties(element),
			extLst: extensionList(element),
		},
	);
}

const SERIES = new Set([
	'idx',
	'order',
	'tx',
	'spPr',
	'invertIfNegative',
	'dPt',
	'dLbls',
	'cat',
	'val',
	'xVal',
	'yVal',
	'smooth',
	'marker',
	'explosion',
	'bubbleSize',
	'bubble3D',
	'shape',
	'extLst',
]);

/** One series (`c:ser`). Trendlines, error bars and picture options are reported, not modelled. */
export function parseSeries(context: ChartParseContext, element: XmlElement): ChartSpaceSeries {
	reportUnmodelled(context, element, SERIES);
	return assignDefined<ChartSpaceSeries>(
		{ dataPoints: children(element, 'dPt', NS.c).map((point) => parseDataPoint(context, point)) },
		{
			index: cNumber(context, element, 'idx'),
			order: cNumber(context, element, 'order'),
			tx: parseChartText(context, cChild(element, 'tx'), true),
			spPr: shapeProperties(element),
			marker: parseMarker(context, cChild(element, 'marker')),
			dataLabels: parseDataLabels(context, cChild(element, 'dLbls')),
			categories: parseDataSource(cChild(element, 'cat')),
			values: parseDataSource(cChild(element, 'val')),
			xValues: parseDataSource(cChild(element, 'xVal')),
			yValues: parseDataSource(cChild(element, 'yVal')),
			bubbleSizes: parseDataSource(cChild(element, 'bubbleSize')),
			smooth: cBool(context, element, 'smooth'),
			invertIfNegative: cBool(context, element, 'invertIfNegative'),
			explosion: cNumber(context, element, 'explosion'),
			bubble3D: cBool(context, element, 'bubble3D'),
			shape: cVal(element, 'shape'),
			extLst: extensionList(element),
		},
	);
}
