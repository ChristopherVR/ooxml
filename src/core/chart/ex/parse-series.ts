// `cx:chartSpace` parser: series (`cx:series`) with their data labels and family-specific layout
// properties, and axes (`cx:axis`).
import type { XmlElement } from '../../xml/index';
import type {
	ChartExBinning,
	ChartExDataLabel,
	ChartExDataLabelOptions,
	ChartExDataLabels,
	ChartExDataPoint,
	ChartExGeography,
	ChartExLabelVisibility,
	ChartExLayoutProperties,
	ChartExSeries,
	ChartExSeriesVisibility,
} from './model';
import {
	assignDefined,
	attribute,
	autoNumberAttribute,
	boolAttribute,
	cxChild,
	cxChildren,
	extensionList,
	indexAttribute,
	numberAttribute,
	numberFormat,
	rawChild,
	reportUnmodelled,
	shapeProperties,
	textProperties,
	valNumber,
	type ChartParseContext,
} from './parse-util';
import { parseData, parseText } from './parse-text';

function labelVisibility(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartExLabelVisibility | undefined {
	if (!element) return undefined;
	return assignDefined<ChartExLabelVisibility>(
		{},
		{
			seriesName: boolAttribute(context, element, 'seriesName'),
			categoryName: boolAttribute(context, element, 'categoryName'),
			value: boolAttribute(context, element, 'value'),
		},
	);
}

type Defined<T> = { [K in keyof T]?: T[K] | undefined };

function labelOptions(
	context: ChartParseContext,
	element: XmlElement,
): Defined<ChartExDataLabelOptions> {
	const separator = cxChild(element, 'separator');
	return {
		position: attribute(element, 'pos'),
		numberFormat: numberFormat(element),
		spPr: shapeProperties(element),
		txPr: textProperties(element),
		visibility: labelVisibility(context, cxChild(element, 'visibility')),
		separator: separator ? (separator.textContent ?? '') : undefined,
		extLst: extensionList(element),
	};
}

const LABEL = new Set(['numFmt', 'spPr', 'txPr', 'visibility', 'separator', 'extLst']);
const LABELS = new Set([...LABEL, 'dataLabel', 'dataLabelHidden']);

function parseDataLabels(context: ChartParseContext, element: XmlElement): ChartExDataLabels {
	reportUnmodelled(context, element, LABELS);
	const labels = cxChildren(element, 'dataLabel').map((label) => {
		reportUnmodelled(context, label, LABEL);
		return assignDefined<ChartExDataLabel>(
			{},
			{ index: indexAttribute(context, label, 'idx'), ...labelOptions(context, label) },
		);
	});
	const hidden = cxChildren(element, 'dataLabelHidden')
		.map((item) => indexAttribute(context, item, 'idx'))
		.filter((index): index is number => index !== undefined);
	return assignDefined<ChartExDataLabels>({ labels, hidden }, labelOptions(context, element));
}

function parseBinning(context: ChartParseContext, element: XmlElement): ChartExBinning {
	reportUnmodelled(context, element, new Set(['binSize', 'binCount']));
	// PowerPoint writes `<cx:binCount val="4"/>`; the text form is read too.
	const scalar = (local: string) => {
		const child = cxChild(element, local);
		if (!child) return undefined;
		if (child.hasAttribute('val')) return numberAttribute(context, child, 'val');
		const value = Number((child.textContent ?? '').trim() || Number.NaN);
		return Number.isFinite(value) ? value : undefined;
	};
	return assignDefined<ChartExBinning>(
		{},
		{
			intervalClosed: attribute(element, 'intervalClosed'),
			underflow: autoNumberAttribute(context, element, 'underflow'),
			overflow: autoNumberAttribute(context, element, 'overflow'),
			binSize: scalar('binSize'),
			binCount: scalar('binCount'),
		},
	);
}

const VISIBILITY = ['connectorLines', 'meanLine', 'meanMarker', 'nonoutliers', 'outliers'] as const;

function parseGeography(context: ChartParseContext, element: XmlElement): ChartExGeography {
	reportUnmodelled(context, element, new Set(['geoCache', 'extLst']));
	return assignDefined<ChartExGeography>(
		{},
		{
			projectionType: attribute(element, 'projectionType'),
			viewedRegionType: attribute(element, 'viewedRegionType'),
			cultureLanguage: attribute(element, 'cultureLanguage'),
			cultureRegion: attribute(element, 'cultureRegion'),
			attribution: attribute(element, 'attribution'),
			geoCacheXml: rawChild(element, 'geoCache'),
			extLst: extensionList(element),
		},
	);
}

const LAYOUT = new Set([
	'parentLabelLayout',
	'regionLabelLayout',
	'visibility',
	'aggregation',
	'binning',
	'geography',
	'statistics',
	'subtotals',
	'extLst',
]);

function parseLayout(context: ChartParseContext, element: XmlElement): ChartExLayoutProperties {
	reportUnmodelled(context, element, LAYOUT);
	const visibility = cxChild(element, 'visibility');
	const binning = cxChild(element, 'binning');
	const geography = cxChild(element, 'geography');
	const statistics = cxChild(element, 'statistics');
	const subtotals = cxChild(element, 'subtotals');
	return assignDefined<ChartExLayoutProperties>(
		{},
		{
			parentLabelLayout: attribute(cxChild(element, 'parentLabelLayout'), 'val'),
			regionLabelLayout: attribute(cxChild(element, 'regionLabelLayout'), 'val'),
			visibility: visibility
				? assignDefined<ChartExSeriesVisibility>(
						{},
						Object.fromEntries(
							VISIBILITY.map((name) => [name, boolAttribute(context, visibility, name)]),
						),
					)
				: undefined,
			aggregation: cxChild(element, 'aggregation') ? true : undefined,
			binning: binning ? parseBinning(context, binning) : undefined,
			geography: geography ? parseGeography(context, geography) : undefined,
			statistics: statistics
				? assignDefined<{ quartileMethod?: string }>(
						{},
						{ quartileMethod: attribute(statistics, 'quartileMethod') },
					)
				: undefined,
			subtotals: subtotals
				? cxChildren(subtotals, 'idx')
						.map((idx) => indexAttribute(context, idx, 'val'))
						.filter((index): index is number => index !== undefined)
				: undefined,
			extLst: extensionList(element),
		},
	);
}

const SERIES = new Set([
	'tx',
	'spPr',
	'txPr',
	'valueColors',
	'valueColorPositions',
	'dataPt',
	'dataLabels',
	'data',
	'dataId',
	'layoutPr',
	'axisId',
	'extLst',
]);

/** One `cx:series`. */
export function parseSeries(context: ChartParseContext, element: XmlElement): ChartExSeries {
	reportUnmodelled(context, element, SERIES);
	const dataPoints = cxChildren(element, 'dataPt').map((point) => {
		reportUnmodelled(context, point, new Set(['spPr', 'extLst']));
		return assignDefined<ChartExDataPoint>(
			{},
			{
				index: indexAttribute(context, point, 'idx'),
				spPr: shapeProperties(point),
				extLst: extensionList(point),
			},
		);
	});
	const labels = cxChild(element, 'dataLabels');
	const inline = cxChild(element, 'data');
	const layout = cxChild(element, 'layoutPr');
	const axisIds = cxChildren(element, 'axisId')
		.map((axis) => indexAttribute(context, axis, 'val'))
		.filter((id): id is number => id !== undefined);
	return assignDefined<ChartExSeries>(
		{ layoutId: attribute(element, 'layoutId') ?? '', dataPoints, axisIds },
		{
			hidden: boolAttribute(context, element, 'hidden'),
			ownerIndex: numberAttribute(context, element, 'ownerIdx'),
			uniqueId: attribute(element, 'uniqueId'),
			formatIndex: numberAttribute(context, element, 'formatIdx'),
			tx: parseText(context, cxChild(element, 'tx')),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			valueColorsXml: rawChild(element, 'valueColors'),
			valueColorPositionsXml: rawChild(element, 'valueColorPositions'),
			dataLabels: labels ? parseDataLabels(context, labels) : undefined,
			data: inline ? parseData(context, inline) : undefined,
			dataId: valNumber(context, element, 'dataId'),
			layoutProperties: layout ? parseLayout(context, layout) : undefined,
			extLst: extensionList(element),
		},
	);
}
