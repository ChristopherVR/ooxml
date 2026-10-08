// `c:chartSpace` parser: the chrome beside the plot (floor and walls, the data table, a value
// axis's display units) and the Office 2017 "show #N/A as blank" chart extension.
import { buildXml, elements, first, parseXml, type XmlElement } from '../xml/index';
import type {
	ChartDataTable,
	ChartDisplayUnits,
	ChartDisplayUnitsLabel,
	ChartSurface,
} from './model-chrome';
import { parseChartText } from './parse-text';
import {
	assignDefined,
	attribute,
	cBool,
	cChild,
	cNumber,
	cVal,
	extensionList,
	manualLayout,
	reportUnmodelled,
	shapeProperties,
	textProperties,
	type ChartParseContext,
} from './parse-util';

const SURFACE = new Set(['thickness', 'spPr', 'pictureOptions', 'extLst']);

/** A floor or wall (`c:floor`, `c:sideWall`, `c:backWall`), when written. */
export function parseSurface(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartSurface | undefined {
	if (!element) return undefined;
	reportUnmodelled(context, element, SURFACE);
	const pictureOptions = cChild(element, 'pictureOptions');
	return assignDefined<ChartSurface>(
		{},
		{
			thickness: cNumber(context, element, 'thickness'),
			spPr: shapeProperties(element),
			pictureOptionsXml: pictureOptions ? buildXml(pictureOptions) : undefined,
			extLst: extensionList(element),
		},
	);
}

const DATA_TABLE = new Set([
	'showHorzBorder',
	'showVertBorder',
	'showOutline',
	'showKeys',
	'spPr',
	'txPr',
	'extLst',
]);

/** The data table (`c:dTable`), when written. */
export function parseDataTable(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartDataTable | undefined {
	if (!element) return undefined;
	reportUnmodelled(context, element, DATA_TABLE);
	return assignDefined<ChartDataTable>(
		{},
		{
			showHorizontalBorder: cBool(context, element, 'showHorzBorder'),
			showVerticalBorder: cBool(context, element, 'showVertBorder'),
			showOutline: cBool(context, element, 'showOutline'),
			showKeys: cBool(context, element, 'showKeys'),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			extLst: extensionList(element),
		},
	);
}

const DISPLAY_UNITS = new Set(['custUnit', 'builtInUnit', 'dispUnitsLbl', 'extLst']);
const DISPLAY_UNITS_LABEL = new Set(['layout', 'tx', 'spPr', 'txPr']);

/** Display units of a value axis (`c:dispUnits`), when written. */
export function parseDisplayUnits(
	context: ChartParseContext,
	element: XmlElement | undefined,
): ChartDisplayUnits | undefined {
	if (!element) return undefined;
	reportUnmodelled(context, element, DISPLAY_UNITS);
	const labelElement = cChild(element, 'dispUnitsLbl');
	let label: ChartDisplayUnitsLabel | undefined;
	if (labelElement) {
		reportUnmodelled(context, labelElement, DISPLAY_UNITS_LABEL);
		label = assignDefined<ChartDisplayUnitsLabel>(
			{},
			{
				layout: manualLayout(labelElement),
				tx: parseChartText(context, cChild(labelElement, 'tx'), false),
				spPr: shapeProperties(labelElement),
				txPr: textProperties(labelElement),
			},
		);
	}
	return assignDefined<ChartDisplayUnits>(
		{},
		{
			builtInUnit: cVal(element, 'builtInUnit'),
			customUnit: cNumber(context, element, 'custUnit'),
			label,
			extLst: extensionList(element),
		},
	);
}

/** `c:ext/@uri` of the Office 2017 chart data display options (`c16r3:dataDisplayOptions16`). */
const DATA_DISPLAY_OPTIONS_URI = '{56B9EC1D-385E-4148-901F-78D8002777C0}';
const C16R3 = 'http://schemas.microsoft.com/office/drawing/2017/03/chart';

/**
 * "Show #N/A as an empty cell" (`c16r3:dispNaAsBlank`), read from the chart's extension list
 * (`ChartSpace.chartExtLst`, which the writer keeps as written). `undefined` when not written.
 * A bare element is true, as for every `CT_Boolean`.
 */
export function readDisplayNaAsBlank(chartExtLst: string | undefined): boolean | undefined {
	if (!chartExtLst?.includes(DATA_DISPLAY_OPTIONS_URI)) return undefined;
	let root: XmlElement;
	try {
		root = parseXml(chartExtLst, { label: 'Chart extension list' }).documentElement;
	} catch {
		return undefined;
	}
	for (const ext of elements(root)) {
		if (ext.localName !== 'ext' || attribute(ext, 'uri') !== DATA_DISPLAY_OPTIONS_URI) continue;
		const flag = first(first(ext, 'dataDisplayOptions16', C16R3), 'dispNaAsBlank', C16R3);
		if (!flag) continue;
		const raw = attribute(flag, 'val')?.trim().toLowerCase();
		return raw === undefined || raw === '1' || raw === 'true';
	}
	return undefined;
}
