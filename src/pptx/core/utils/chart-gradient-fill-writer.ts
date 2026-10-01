/**
 * Write-side mirror of `chart-gradient-fill.ts`: turns a modelled
 * {@link PptxChartGradientFill} back into a DrawingML `a:gradFill` inside a
 * chart series' or data point's `c:spPr`.
 *
 * ## Lossless contract
 *
 * The chart part is re-parsed from the archive on save and mutated in place,
 * so the authored `a:gradFill` is still there to compare against.
 * {@link applyChartGradientToSpPr} re-parses it with the same codec the load
 * used and leaves it byte-for-byte alone when the model still describes it
 * (so a themed `a:schemeClr` stop with `lumMod`/`lumOff` never collapses to
 * an `a:srgbClr`). Only an edited, added or removed gradient touches the
 * node, and an edit re-emits each unchanged stop's authored colour choice.
 *
 * `a:gradFill` is a member of the `EG_FillProperties` choice, so writing one
 * removes every other member (`a:noFill`, `a:solidFill`, `a:pattFill`, ...)
 * and lands before `a:ln` / `a:effectLst` / `a:scene3d` / `a:sp3d`, as
 * `CT_ShapeProperties` requires.
 *
 * Line-drawn families (line/line3D/scatter/radar/stock) have no fillable
 * area: a gradient on one of their series is ignored here (never written,
 * never an error); the SDK setters reject it up front instead.
 *
 * @module utils/chart-gradient-fill-writer
 */
import type { PptxChartGradientFill, PptxChartType, XmlObject } from '../types';
import type { ResolveChartColor } from './chart-color-choice';
import { buildChartGradFillXml, chartGradientsEqual } from './chart-gradient-fill-xml';

export type { AuthoredChartGradient } from './chart-gradient-fill-xml';
export {
	buildChartGradFillXml,
	chartGradientsEqual,
	DEFAULT_CHART_GRADIENT_ANGLE,
} from './chart-gradient-fill-xml';

type GetLocalName = (key: string) => string;

/** Re-parse an authored `c:spPr`'s `a:gradFill` the way the load did. */
export type ParseChartGradient = (spPr: XmlObject | undefined) => PptxChartGradientFill | undefined;

/** What a save-side caller knows about the chart part being written. */
export interface ChartGradientWriteOptions {
	/** The load-time parser; omitted for a freshly generated chart (nothing to preserve). */
	parseGradient?: ParseChartGradient;
	/** Resolves an authored stop colour, so an unchanged stop keeps its theme colour. */
	resolveColor?: ResolveChartColor;
	/** Line-drawn series/point: gradients are ignored (see the module docblock). */
	lineDrawn?: boolean;
}

const FILL_CHOICE = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
/** `CT_ShapeProperties` children that follow the fill choice. */
const AFTER_FILL = new Set(['ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst']);

/** Chart types whose series paint an area, so a direct `c:spPr` fill applies. */
const AREA_FILLED = new Set<PptxChartType>([
	'bar',
	'bar3D',
	'area',
	'area3D',
	'pie',
	'pie3D',
	'doughnut',
	'ofPie',
	'bubble',
	'surface',
	'combo',
]);

/** Whether a series drawn as `type` can carry a `c:spPr/a:gradFill`. */
export function supportsChartGradientFill(type: PptxChartType | undefined): boolean {
	return type !== undefined && AREA_FILLED.has(type);
}

function findKey(obj: XmlObject, local: string, getLocalName: GetLocalName): string | undefined {
	return Object.keys(obj).find((k) => getLocalName(k) === local);
}

/**
 * Replace whatever `EG_FillProperties` member `spPr` has with `a:gradFill`,
 * keeping `CT_ShapeProperties` order (fill before `a:ln` and the effects).
 */
function setGradFill(spPr: XmlObject, gradFill: XmlObject, getLocalName: GetLocalName): void {
	const entries = Object.entries(spPr).filter(([key]) => !FILL_CHOICE.has(getLocalName(key)));
	const at = entries.findIndex(([key]) => AFTER_FILL.has(getLocalName(key)));
	entries.splice(at === -1 ? entries.length : at, 0, ['a:gradFill', gradFill]);
	for (const key of Object.keys(spPr)) {
		delete spPr[key];
	}
	for (const [key, value] of entries) {
		spPr[key] = value;
	}
}

/**
 * Reconcile a `c:spPr`'s fill with the modelled gradient. Mutates `spPr` and
 * returns whether it changed.
 *
 * - `gradient` matches the authored `a:gradFill`: untouched (lossless).
 * - `gradient` set otherwise: written as the one fill choice.
 * - `gradient` absent but a parseable `a:gradFill` authored: the model
 *   dropped it, so it is removed (a fill the parser could not read, and so
 *   never modelled, is left alone).
 * - `options.lineDrawn`: nothing is touched.
 */
export function applyChartGradientToSpPr(
	spPr: XmlObject,
	gradient: PptxChartGradientFill | undefined,
	getLocalName: GetLocalName,
	options: ChartGradientWriteOptions = {},
): boolean {
	if (options.lineDrawn) {
		return false;
	}
	const authoredKey = findKey(spPr, 'gradFill', getLocalName);
	const authoredNode = authoredKey ? (spPr[authoredKey] as XmlObject) : undefined;
	const parsed = authoredNode ? options.parseGradient?.(spPr) : undefined;
	if (!gradient) {
		if (authoredKey && parsed) {
			delete spPr[authoredKey];
			return true;
		}
		return false;
	}
	if (parsed && chartGradientsEqual(parsed, gradient)) {
		return false;
	}
	const authored = authoredNode && parsed ? { node: authoredNode, parsed } : undefined;
	setGradFill(
		spPr,
		buildChartGradFillXml(gradient, authored, getLocalName, options.resolveColor),
		getLocalName,
	);
	return true;
}
