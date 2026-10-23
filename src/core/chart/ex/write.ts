// `cx:chartSpace` writer: one serializer for the neutral extended chart model, the counterpart of
// `parseChartExSpace`. Elements follow the chartex schema's sequences; booleans are written `1`/`0`,
// numbers in the canonical form of `chartNumber`, cached point text exactly as held. Raw pieces
// (extension lists, geography caches, colour scales, format overrides, print settings) are written
// verbatim minus the declarations the root already makes. What the parser reports as not modelled
// is not written.
import { NS } from '../../xml/index';
import { escapeAttribute } from '../write-util';
import type { ChartExSpace } from './model';
import { axisXml, chartDataXml, seriesXml, titleXml } from './write-series';
import { attributes, cxElement, raw, shapeXml, textBodyXml } from './write-util';

/** Options of {@link writeChartExSpace}. */
export interface WriteChartExSpaceOptions {
	/** What follows the XML declaration: Office writes nothing in chartex parts (the default). */
	declarationBreak?: '' | '\n' | '\r\n';
	/**
	 * Write empty elements as `<cx:x />` (with a space), as the .NET producers behind Word's and
	 * Excel's chartex parts do; the default is `<cx:x/>`.
	 */
	emptyElementSpace?: boolean;
}

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The declarations a chartex root makes when the model carries none. */
export const CHART_EX_ROOT_DECLARATIONS: readonly { prefix: string; uri: string }[] = [
	{ prefix: 'cx', uri: NS.cx },
	{ prefix: 'a', uri: NS.a },
	{ prefix: 'r', uri: NS.r },
];

/** The writer options that reproduce how `xml` was written (declaration break, empty elements). */
export function chartExWriteOptionsOf(xml: string): WriteChartExSpaceOptions {
	const declarationBreak = (/^<\?xml[^>]*\?>(\r\n|\n)?/.exec(xml)?.[1] ?? '') as '' | '\n' | '\r\n';
	return { declarationBreak, ...(/<[^>]*\s\/>/.test(xml) ? { emptyElementSpace: true } : {}) };
}

/**
 * Writes an extended chart part (`cx:chartSpace`) from the neutral model. The root declares the
 * model's namespaces in their order, plus `cx`, `a` and `r` (and `c` for user shapes) when missing.
 */
export function writeChartExSpace(
	space: ChartExSpace,
	options: WriteChartExSpaceOptions = {},
): string {
	const bindings = new Map<string, string>();
	for (const { prefix, uri } of space.namespaceDeclarations ?? CHART_EX_ROOT_DECLARATIONS)
		if (!bindings.has(prefix) && /^[A-Za-z_][\w.-]*$/.test(prefix)) bindings.set(prefix, uri);
	for (const { prefix, uri } of CHART_EX_ROOT_DECLARATIONS)
		if (!bindings.has(prefix)) bindings.set(prefix, uri);
	if (space.userShapesRelId !== undefined && !bindings.has('c')) bindings.set('c', NS.c);
	const context = { bindings };
	const declarations = [...bindings]
		.map(([prefix, uri]) => ` xmlns:${prefix}="${escapeAttribute(uri)}"`)
		.join('');
	const plot = space.plotArea;
	const region = cxElement(
		'plotAreaRegion',
		(plot.plotSurface
			? cxElement(
					'plotSurface',
					shapeXml(context, plot.plotSurface.spPr) + raw(context, plot.plotSurface.extLst),
				)
			: '') +
			plot.series.map((series) => seriesXml(context, series)).join('') +
			raw(context, plot.regionExtLst),
	);
	const plotArea = cxElement(
		'plotArea',
		region +
			plot.axes.map((axis) => axisXml(context, axis)).join('') +
			shapeXml(context, plot.spPr) +
			raw(context, plot.extLst),
	);
	const legend = space.legend;
	const legendXml = legend
		? cxElement(
				'legend',
				shapeXml(context, legend.spPr) +
					textBodyXml(context, legend.txPr) +
					raw(context, legend.extLst),
				attributes([
					['pos', legend.position],
					['align', legend.align],
					['overlay', legend.overlay],
				]),
			)
		: '';
	const chart = cxElement(
		'chart',
		titleXml(context, space.title) + plotArea + legendXml + raw(context, space.chartExtLst),
	);
	const userShapes =
		space.userShapesRelId === undefined
			? ''
			: `<c:userShapes${attributes([['r:id', space.userShapesRelId]])}/>`;
	const body =
		chartDataXml(context, space.chartData) +
		chart +
		shapeXml(context, space.spPr) +
		textBodyXml(context, space.txPr) +
		raw(context, space.colorMapOverrideXml) +
		raw(context, space.formatOverridesXml) +
		raw(context, space.printSettingsXml) +
		userShapes +
		raw(context, space.extLst);
	let xml = `${DECLARATION}${options.declarationBreak ?? ''}<cx:chartSpace${declarations}>${body}</cx:chartSpace>`;
	// Attribute values and text escape `>`, so every `/>` ends an empty element.
	if (options.emptyElementSpace) xml = xml.replace(/\/>/g, ' />');
	return xml;
}
