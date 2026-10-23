// `cx:chartSpace` parser: chart data (`cx:chartData`), chart text, titles and the legend.
import { NS, attr, elements, relAttr, type XmlElement } from '../../xml/index';
import type { ChartCachePoint } from '../model-series';
import { chartTextBody } from '../parse-util';
import type {
	ChartExChartData,
	ChartExData,
	ChartExDimension,
	ChartExLegend,
	ChartExLevel,
	ChartExText,
	ChartExTitle,
} from './model';
import {
	assignDefined,
	attribute,
	boolAttribute,
	cxChild,
	cxChildren,
	extensionList,
	formula,
	indexAttribute,
	reportUnmodelled,
	shapeProperties,
	textProperties,
	type ChartParseContext,
} from './parse-util';

const LEVEL = new Set(['pt', 'extLst']);

function parseLevel(context: ChartParseContext, element: XmlElement): ChartExLevel {
	reportUnmodelled(context, element, LEVEL);
	const points = cxChildren(element, 'pt').map((pt) => {
		const point: ChartCachePoint = { value: pt.textContent ?? '' };
		const index = indexAttribute(context, pt, 'idx');
		if (index !== undefined) point.index = index;
		return point;
	});
	return assignDefined<ChartExLevel>(
		{ points },
		{
			pointCount: indexAttribute(context, element, 'ptCount'),
			formatCode: attribute(element, 'formatCode'),
			name: attribute(element, 'name'),
			extLst: extensionList(element),
		},
	);
}

const DIMENSION = new Set(['f', 'nf', 'lvl', 'extLst']);

function parseDimension(context: ChartParseContext, element: XmlElement): ChartExDimension {
	reportUnmodelled(context, element, DIMENSION);
	return assignDefined<ChartExDimension>(
		{
			kind: element.localName as ChartExDimension['kind'],
			levels: cxChildren(element, 'lvl').map((level) => parseLevel(context, level)),
		},
		{
			type: attribute(element, 'type'),
			formula: formula(element, 'f'),
			nameFormula: formula(element, 'nf'),
			extLst: extensionList(element),
		},
	);
}

const DATA = new Set(['numDim', 'strDim', 'extLst']);

/** One `cx:data` element. */
export function parseData(context: ChartParseContext, element: XmlElement): ChartExData {
	reportUnmodelled(context, element, DATA);
	const dimensions = elements(element)
		.filter((child) => child.localName === 'numDim' || child.localName === 'strDim')
		.map((child) => parseDimension(context, child));
	return assignDefined<ChartExData>(
		{ dimensions },
		{ id: indexAttribute(context, element, 'id'), extLst: extensionList(element) },
	);
}

const CHART_DATA = new Set(['externalData', 'data']);

/** `cx:chartData`; an absent element reads as no data. */
export function parseChartData(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartExChartData {
	if (!element) return { data: [] };
	reportUnmodelled(context, element, CHART_DATA);
	const out: ChartExChartData = {
		data: cxChildren(element, 'data').map((data) => parseData(context, data)),
	};
	const external = cxChild(element, 'externalData');
	if (external) {
		const autoUpdate = attr(external, 'autoUpdate', NS.cx, 'cx') ?? undefined;
		out.externalData = assignDefined<NonNullable<ChartExChartData['externalData']>>(
			{},
			{
				relId: relAttr(external, 'id') || undefined,
				autoUpdate:
					autoUpdate === undefined
						? undefined
						: autoUpdate === '1' || autoUpdate === 'true'
							? true
							: autoUpdate === '0' || autoUpdate === 'false'
								? false
								: undefined,
			},
		);
	}
	return out;
}

/** Chart text (`cx:tx`): `cx:txData` (formula and cached value) or `cx:rich`. */
export function parseText(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartExText | undefined {
	if (!element) return undefined;
	reportUnmodelled(context, element, new Set(['txData', 'rich']));
	const data = cxChild(element, 'txData');
	if (data) reportUnmodelled(context, data, new Set(['f', 'v']));
	const value = cxChild(data, 'v');
	return assignDefined<ChartExText>(
		{},
		{
			formula: formula(data, 'f'),
			value: value ? (value.textContent ?? '') : undefined,
			rich: chartTextBody(cxChild(element, 'rich')),
		},
	);
}

const TITLE = new Set(['tx', 'spPr', 'txPr', 'extLst']);

/** A title (`cx:title`, `cx:unitsLabel`). */
export function parseTitle(context: ChartParseContext, element: XmlElement): ChartExTitle {
	reportUnmodelled(context, element, TITLE);
	return assignDefined<ChartExTitle>(
		{},
		{
			tx: parseText(context, cxChild(element, 'tx')),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			position: attribute(element, 'pos'),
			align: attribute(element, 'align'),
			overlay: boolAttribute(context, element, 'overlay'),
			extLst: extensionList(element),
		},
	);
}

const LEGEND = new Set(['spPr', 'txPr', 'extLst']);

/** `cx:legend`. */
export function parseLegend(context: ChartParseContext, element: XmlElement): ChartExLegend {
	reportUnmodelled(context, element, LEGEND);
	return assignDefined<ChartExLegend>(
		{},
		{
			position: attribute(element, 'pos'),
			align: attribute(element, 'align'),
			overlay: boolAttribute(context, element, 'overlay'),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			extLst: extensionList(element),
		},
	);
}
