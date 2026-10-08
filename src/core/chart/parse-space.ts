// `c:chartSpace` parser over the shared `xml` DOM: one reader for every Office format. It never
// throws on content it does not know; such children are reported as issues and stay in the part.
import { NS, buildXml, children, first, parseXml, relAttr, type XmlElement } from '../xml/index';
import type { ChartLegend, ChartSpace, ChartSpaceParseResult, ChartView3D } from './model';
import { parseSurface } from './parse-chrome';
import { parsePlotArea } from './parse-plot';
import { parseTitle } from './parse-text';
import { CHART_ROOT_BINDINGS } from './xml-fragment';
import {
	attribute,
	assignDefined,
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

const LEGEND = new Set(['legendPos', 'legendEntry', 'layout', 'overlay', 'spPr', 'txPr', 'extLst']);

function parseLegend(context: ChartParseContext, element: XmlElement): ChartLegend {
	reportUnmodelled(context, element, LEGEND);
	const entries = children(element, 'legendEntry', NS.c).map((entry) =>
		assignDefined<ChartLegend['entries'][number]>(
			{},
			{
				index: cNumber(context, entry, 'idx'),
				deleted: cBool(context, entry, 'delete'),
				txPr: textProperties(entry),
			},
		),
	);
	return assignDefined<ChartLegend>(
		{ entries },
		{
			position: cVal(element, 'legendPos'),
			overlay: cBool(context, element, 'overlay'),
			layout: manualLayout(element),
			spPr: shapeProperties(element),
			txPr: textProperties(element),
			extLst: extensionList(element),
		},
	);
}

const VIEW_3D = new Set(['rotX', 'rotY', 'depthPercent', 'hPercent', 'rAngAx', 'perspective']);

function parseView3D(context: ChartParseContext, element: XmlElement): ChartView3D {
	reportUnmodelled(context, element, VIEW_3D);
	return assignDefined<ChartView3D>(
		{},
		{
			rotX: cNumber(context, element, 'rotX'),
			rotY: cNumber(context, element, 'rotY'),
			depthPercent: cNumber(context, element, 'depthPercent'),
			heightPercent: cNumber(context, element, 'hPercent'),
			rightAngleAxes: cBool(context, element, 'rAngAx'),
			perspective: cNumber(context, element, 'perspective'),
		},
	);
}

/** `c:style`, directly or as the fallback of Office's `mc:AlternateContent` (`c14:style`). */
function chartStyle(context: ChartParseContext, space: XmlElement): number | undefined {
	const direct = cNumber(context, space, 'style');
	if (direct !== undefined) return direct;
	for (const alternate of children(space, 'AlternateContent', NS.mc)) {
		const style = cNumber(context, first(alternate, 'Fallback', NS.mc), 'style');
		if (style !== undefined) return style;
	}
	return undefined;
}

/** `c14:style` of the `mc:Choice` Office writes beside the `c:style` fallback. */
function c14Style(space: XmlElement): number | undefined {
	for (const alternate of children(space, 'AlternateContent', NS.mc)) {
		const raw = attribute(first(first(alternate, 'Choice', NS.mc), 'style', NS.c14, false), 'val');
		const value = raw?.trim() ? Number(raw) : Number.NaN;
		if (Number.isFinite(value)) return value;
	}
	return undefined;
}

/** Root namespace declarations other than the `c`, `a` and `r` every chart part declares. */
function namespaceDeclarations(root: XmlElement): { prefix: string; uri: string }[] | undefined {
	const out: { prefix: string; uri: string }[] = [];
	for (const item of Array.from(root.attributes)) {
		if (item.prefix !== 'xmlns' || CHART_ROOT_BINDINGS.get(item.localName) === item.value) continue;
		out.push({ prefix: item.localName, uri: item.value });
	}
	return out.length ? out : undefined;
}

const CHART = new Set([
	'title',
	'autoTitleDeleted',
	'view3D',
	'floor',
	'sideWall',
	'backWall',
	'plotArea',
	'legend',
	'plotVisOnly',
	'dispBlanksAs',
	'showDLblsOverMax',
	'extLst',
]);
const SPACE = new Set([
	'date1904',
	'lang',
	'roundedCorners',
	'AlternateContent',
	'style',
	'chart',
	'spPr',
	'txPr',
	'externalData',
	'printSettings',
	'userShapes',
	'extLst',
]);

/**
 * Parses a chart part (`c:chartSpace`) from its XML text or root element. Anything the model does
 * not cover is listed in `issues`; the caller keeps the part itself for round-trip.
 */
export function parseChartSpace(source: string | XmlElement): ChartSpaceParseResult {
	const context: ChartParseContext = { issues: [] };
	const root =
		typeof source === 'string' ? parseXml(source, { label: 'Chart part' }).documentElement : source;
	if (root.localName !== 'chartSpace' || (root.namespaceURI && root.namespaceURI !== NS.c)) {
		context.issues.push({
			code: 'CHART_ROOT_UNEXPECTED',
			message: `The part root is {${root.namespaceURI ?? ''}}${root.localName}, not c:chartSpace (chartex parts are not modelled).`,
		});
		return { chartSpace: { plotArea: { groups: [], axes: [] } }, issues: context.issues };
	}
	reportUnmodelled(context, root, SPACE);
	const chart = cChild(root, 'chart');
	if (!chart)
		context.issues.push({ code: 'CHART_CHART_MISSING', message: 'The part has no c:chart.' });
	else reportUnmodelled(context, chart, CHART);
	const title = cChild(chart, 'title');
	const legend = cChild(chart, 'legend');
	const view3D = cChild(chart, 'view3D');
	const printSettings = cChild(root, 'printSettings');
	const chartSpace = assignDefined<ChartSpace>(
		{ plotArea: parsePlotArea(context, cChild(chart, 'plotArea')) },
		{
			date1904: cBool(context, root, 'date1904'),
			language: cVal(root, 'lang'),
			roundedCorners: cBool(context, root, 'roundedCorners'),
			style: chartStyle(context, root),
			c14Style: c14Style(root),
			title: title ? parseTitle(context, title) : undefined,
			autoTitleDeleted: cBool(context, chart, 'autoTitleDeleted'),
			view3D: view3D ? parseView3D(context, view3D) : undefined,
			floor: parseSurface(context, cChild(chart, 'floor')),
			sideWall: parseSurface(context, cChild(chart, 'sideWall')),
			backWall: parseSurface(context, cChild(chart, 'backWall')),
			legend: legend ? parseLegend(context, legend) : undefined,
			plotVisibleOnly: cBool(context, chart, 'plotVisOnly'),
			displayBlanksAs: cVal(chart, 'dispBlanksAs'),
			showDataLabelsOverMax: cBool(context, chart, 'showDLblsOverMax'),
			spPr: shapeProperties(root),
			txPr: textProperties(root),
			externalDataRelId: relAttr(cChild(root, 'externalData'), 'id') || undefined,
			externalDataAutoUpdate: cBool(context, cChild(root, 'externalData'), 'autoUpdate'),
			printSettings: printSettings ? buildXml(printSettings) : undefined,
			userShapesRelId: relAttr(cChild(root, 'userShapes'), 'id') || undefined,
			chartExtLst: extensionList(chart),
			extLst: extensionList(root),
			namespaceDeclarations: namespaceDeclarations(root),
		},
	);
	return { chartSpace, issues: context.issues };
}
