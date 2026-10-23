// `cx:chartSpace` parser: axes (`cx:axis`) with their scaling, title, display units, gridlines,
// tick marks and labels.
import type { XmlElement } from '../../xml/index';
import type { ChartExAxis, ChartExAxisScaling, ChartExAxisUnits } from './model';
import {
	assignDefined,
	attribute,
	autoNumberAttribute,
	boolAttribute,
	cxChild,
	extensionList,
	indexAttribute,
	lines,
	numberFormat,
	reportUnmodelled,
	shapeProperties,
	textProperties,
	type ChartParseContext,
} from './parse-util';
import { parseTitle } from './parse-text';

function parseScaling(
	context: ChartParseContext,
	axis: XmlElement,
): ChartExAxisScaling | undefined {
	const category = cxChild(axis, 'catScaling');
	if (category)
		return assignDefined<ChartExAxisScaling>(
			{ kind: 'cat' },
			{ gapWidth: autoNumberAttribute(context, category, 'gapWidth') },
		);
	const value = cxChild(axis, 'valScaling');
	if (!value) return undefined;
	return assignDefined<ChartExAxisScaling & { kind: 'val' }>(
		{ kind: 'val' },
		{
			max: autoNumberAttribute(context, value, 'max'),
			min: autoNumberAttribute(context, value, 'min'),
			majorUnit: autoNumberAttribute(context, value, 'majorUnit'),
			minorUnit: autoNumberAttribute(context, value, 'minorUnit'),
		},
	);
}

function parseUnits(context: ChartParseContext, element: XmlElement): ChartExAxisUnits {
	reportUnmodelled(context, element, new Set(['unitsLabel', 'extLst']));
	const label = cxChild(element, 'unitsLabel');
	return assignDefined<ChartExAxisUnits>(
		{},
		{
			unit: attribute(element, 'unit'),
			label: label ? parseTitle(context, label) : undefined,
			extLst: extensionList(element),
		},
	);
}

const AXIS = new Set([
	'catScaling',
	'valScaling',
	'title',
	'units',
	'majorGridlines',
	'minorGridlines',
	'majorTickMarks',
	'minorTickMarks',
	'tickLabels',
	'numFmt',
	'spPr',
	'txPr',
	'extLst',
]);

const tickMarks = (element: XmlElement | undefined) =>
	element ? assignDefined<{ type?: string }>({}, { type: attribute(element, 'type') }) : undefined;

/** One `cx:axis`. */
export function parseAxis(context: ChartParseContext, element: XmlElement): ChartExAxis {
	reportUnmodelled(context, element, AXIS);
	const title = cxChild(element, 'title');
	const units = cxChild(element, 'units');
	const tickLabels = cxChild(element, 'tickLabels');
	return assignDefined<ChartExAxis>(
		{},
		{
			id: indexAttribute(context, element, 'id'),
			hidden: boolAttribute(context, element, 'hidden'),
			scaling: parseScaling(context, element),
			title: title ? parseTitle(context, title) : undefined,
			units: units ? parseUnits(context, units) : undefined,
			majorGridlines: lines(element, 'majorGridlines'),
			minorGridlines: lines(element, 'minorGridlines'),
			majorTickMarks: tickMarks(cxChild(element, 'majorTickMarks')),
			minorTickMarks: tickMarks(cxChild(element, 'minorTickMarks')),
			tickLabels: tickLabels
				? assignDefined<{ extLst?: string }>({}, { extLst: extensionList(tickLabels) })
				: undefined,
			numberFormat: numberFormat(element),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			extLst: extensionList(element),
		},
	);
}
