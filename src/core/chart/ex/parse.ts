// `cx:chartSpace` parser over the shared `xml` DOM: one reader of Office 2016 extended chart parts
// for every Office format. It never throws on content it does not know; such children are reported
// as issues (and are not written back by `writeChartExSpace`).
import { NS, first, parseXml, relAttr, type XmlElement } from '../../xml/index';
import type { ChartExPlotArea, ChartExSpace, ChartExSpaceParseResult } from './model';
import {
	assignDefined,
	cxChild,
	cxChildren,
	extensionList,
	rawChild,
	reportUnmodelled,
	shapeProperties,
	textProperties,
	type ChartParseContext,
} from './parse-util';
import { parseAxis } from './parse-axis';
import { parseSeries } from './parse-series';
import { parseChartData, parseLegend, parseTitle } from './parse-text';

const REGION = new Set(['plotSurface', 'series', 'extLst']);
const PLOT_AREA = new Set(['plotAreaRegion', 'axis', 'spPr', 'extLst']);

function parsePlotArea(context: ChartParseContext, element: XmlElement | undefined) {
	const out: ChartExPlotArea = { series: [], axes: [] };
	if (!element) return out;
	reportUnmodelled(context, element, PLOT_AREA);
	const region = cxChild(element, 'plotAreaRegion');
	if (region) {
		reportUnmodelled(context, region, REGION);
		out.series = cxChildren(region, 'series').map((series) => parseSeries(context, series));
		const surface = cxChild(region, 'plotSurface');
		if (surface) {
			reportUnmodelled(context, surface, new Set(['spPr', 'extLst']));
			out.plotSurface = assignDefined<NonNullable<ChartExPlotArea['plotSurface']>>(
				{},
				{ spPr: shapeProperties(surface), extLst: extensionList(surface) },
			);
		}
		const regionExtLst = extensionList(region);
		if (regionExtLst) out.regionExtLst = regionExtLst;
	}
	out.axes = cxChildren(element, 'axis').map((axis) => parseAxis(context, axis));
	return assignDefined(out, { spPr: shapeProperties(element), extLst: extensionList(element) });
}

/** Every namespace declaration of the root, in document order. */
function namespaceDeclarations(root: XmlElement): { prefix: string; uri: string }[] {
	return Array.from(root.attributes)
		.filter((item) => item.prefix === 'xmlns')
		.map((item) => ({ prefix: item.localName, uri: item.value }));
}

const CHART = new Set(['title', 'plotArea', 'legend', 'extLst']);
const SPACE = new Set([
	'chartData',
	'chart',
	'spPr',
	'txPr',
	'clrMapOvr',
	'fmtOvrs',
	'printSettings',
	'userShapes',
	'extLst',
]);

/** Whether a part root is an extended chart (`cx:chartSpace`). */
export const isChartExRoot = (root: XmlElement): boolean =>
	root.localName === 'chartSpace' && root.namespaceURI === NS.cx;

/**
 * Parses an extended chart part (`cx:chartSpace`) from its XML text or root element. Anything the
 * model does not cover is listed in `issues`; the caller keeps the part itself for round-trip.
 */
export function parseChartExSpace(source: string | XmlElement): ChartExSpaceParseResult {
	const context: ChartParseContext = { issues: [] };
	const root =
		typeof source === 'string' ? parseXml(source, { label: 'Chart part' }).documentElement : source;
	if (!isChartExRoot(root)) {
		context.issues.push({
			code: 'CHART_ROOT_UNEXPECTED',
			message: `The part root is {${root.namespaceURI ?? ''}}${root.localName}, not cx:chartSpace (classic parts are read by parseChartSpace).`,
		});
		return {
			chartExSpace: { chartData: { data: [] }, plotArea: { series: [], axes: [] } },
			issues: context.issues,
		};
	}
	reportUnmodelled(context, root, SPACE);
	const chart = cxChild(root, 'chart');
	if (!chart)
		context.issues.push({ code: 'CHART_CHART_MISSING', message: 'The part has no cx:chart.' });
	else reportUnmodelled(context, chart, CHART);
	const title = cxChild(chart, 'title');
	const legend = cxChild(chart, 'legend');
	const userShapes = first(root, 'userShapes', NS.c, false);
	const chartExSpace = assignDefined<ChartExSpace>(
		{
			chartData: parseChartData(context, cxChild(root, 'chartData')),
			plotArea: parsePlotArea(context, cxChild(chart, 'plotArea')),
		},
		{
			title: title ? parseTitle(context, title) : undefined,
			legend: legend ? parseLegend(context, legend) : undefined,
			chartExtLst: extensionList(chart),
			spPr: shapeProperties(root),
			txPr: textProperties(root),
			colorMapOverrideXml: rawChild(root, 'clrMapOvr'),
			formatOverridesXml: rawChild(root, 'fmtOvrs'),
			printSettingsXml: rawChild(root, 'printSettings'),
			userShapesRelId: relAttr(userShapes, 'id') || undefined,
			extLst: extensionList(root),
			namespaceDeclarations: namespaceDeclarations(root),
		},
	);
	return { chartExSpace, issues: context.issues };
}
