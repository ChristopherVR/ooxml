// `c:chartSpace` writer: one serializer for the neutral chart model, the counterpart of
// `parseChartSpace`. Elements follow the sequences of ECMA-376 Part 1, 21.2 (Office rejects parts
// out of schema order); booleans are written `1`/`0`, numbers in the canonical form of
// `chartNumber`, cached point text exactly as held. Raw extension lists, print settings and kept
// sources are written verbatim minus the declarations the root already makes. What the parser
// reports as not modelled (trendlines, error bars, walls, floors, data tables...) is not written.
import { NS } from '../xml/index';
import type { ChartLegend, ChartSpace, ChartView3D } from './model';
import { plotAreaXml } from './write-plot';
import { layoutXml, shapePropertiesXml, textBodyXml } from './write-shape';
import { titleXml } from './write-text';
import { elementXml, escapeAttribute, raw, valXml, type ChartWriteContext } from './write-util';
import { CHART_ROOT_BINDINGS } from './xml-fragment';

/** Options of {@link writeChartSpace}. */
export interface WriteChartSpaceOptions {
	/**
	 * What follows the XML declaration: Excel and PowerPoint write `\r\n` (the default); other
	 * producers write `\n` or nothing.
	 */
	declarationBreak?: '' | '\n' | '\r\n';
}

const DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

function view3DXml(view: ChartView3D | undefined): string {
	if (!view) return '';
	return elementXml(
		'view3D',
		valXml('rotX', view.rotX) +
			valXml('hPercent', view.heightPercent) +
			valXml('rotY', view.rotY) +
			valXml('depthPercent', view.depthPercent) +
			valXml('rAngAx', view.rightAngleAxes) +
			valXml('perspective', view.perspective),
	);
}

function legendXml(context: ChartWriteContext, legend: ChartLegend | undefined): string {
	if (!legend) return '';
	const entries = legend.entries
		.map((entry) =>
			elementXml(
				'legendEntry',
				valXml('idx', entry.index) +
					valXml('delete', entry.deleted) +
					textBodyXml(context, entry.txPr),
			),
		)
		.join('');
	return elementXml(
		'legend',
		valXml('legendPos', legend.position) +
			entries +
			layoutXml(context, legend.layout) +
			valXml('overlay', legend.overlay) +
			shapePropertiesXml(context, legend.spPr) +
			textBodyXml(context, legend.txPr) +
			raw(context, legend.extLst),
	);
}

/** `c:style`, wrapped the way Office writes it when the model has the `c14:style` choice. */
function styleXml(context: ChartWriteContext, space: ChartSpace): string {
	const style = valXml('style', space.style);
	if (space.c14Style === undefined) return style;
	const choice = `<mc:Choice Requires="c14" xmlns:c14="${NS.c14}"><c14:style val="${space.c14Style}"/></mc:Choice>`;
	const fallback = style ? `<mc:Fallback>${style}</mc:Fallback>` : '<mc:Fallback/>';
	return raw(
		context,
		`<mc:AlternateContent xmlns:mc="${NS.mc}">${choice}${fallback}</mc:AlternateContent>`,
	);
}

const relationship = (local: string, id: string | undefined, content = '') =>
	id === undefined ? '' : elementXml(local, content, ` r:id="${escapeAttribute(id)}"`);

/**
 * Writes a chart part (`c:chartSpace`) from the neutral model. The root declares `c`, `a`, `r` and
 * the model's extra declarations (`namespaceDeclarations`); nothing else is added.
 */
export function writeChartSpace(space: ChartSpace, options: WriteChartSpaceOptions = {}): string {
	const bindings = new Map(CHART_ROOT_BINDINGS);
	for (const { prefix, uri } of space.namespaceDeclarations ?? [])
		if (!bindings.has(prefix) && /^[A-Za-z_][\w.-]*$/.test(prefix)) bindings.set(prefix, uri);
	const context: ChartWriteContext = { bindings };
	const declarations = [...bindings]
		.map(([prefix, uri]) => ` xmlns:${prefix}="${escapeAttribute(uri)}"`)
		.join('');
	const chart = elementXml(
		'chart',
		titleXml(context, space.title) +
			valXml('autoTitleDeleted', space.autoTitleDeleted) +
			view3DXml(space.view3D) +
			plotAreaXml(context, space.plotArea) +
			legendXml(context, space.legend) +
			valXml('plotVisOnly', space.plotVisibleOnly) +
			valXml('dispBlanksAs', space.displayBlanksAs) +
			valXml('showDLblsOverMax', space.showDataLabelsOverMax) +
			raw(context, space.chartExtLst),
	);
	const body =
		valXml('date1904', space.date1904) +
		valXml('lang', space.language) +
		valXml('roundedCorners', space.roundedCorners) +
		styleXml(context, space) +
		chart +
		shapePropertiesXml(context, space.spPr) +
		textBodyXml(context, space.txPr) +
		relationship(
			'externalData',
			space.externalDataRelId,
			valXml('autoUpdate', space.externalDataAutoUpdate),
		) +
		raw(context, space.printSettings) +
		relationship('userShapes', space.userShapesRelId) +
		raw(context, space.extLst);
	const breakAfter = options.declarationBreak ?? '\r\n';
	return `${DECLARATION}${breakAfter}<c:chartSpace${declarations}>${body}</c:chartSpace>`;
}
