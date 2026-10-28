// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Reads a `wps:wsp` drawing shape: preset geometry, fill, outline and text. The geometry, fill and
// outline readers are the shared DrawingML ones.
import { parseDrawingFill, parseDrawingGeometry, parseDrawingLine } from '../drawingml/index';
import type { DrawingLine } from '../drawingml/index';
import { buildXml } from '../xml/index';
import { lookupPresetShape } from '../geometry/preset-shape-evaluator';
import type { ShapeSpec } from './shape-model';
import { isElement, type XmlElement } from './xml';

export const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
export const WPS_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function child(parent: XmlElement | undefined, ns: string, local: string): XmlElement | undefined {
	if (!parent) return undefined;
	for (const node of Array.from(parent.childNodes))
		if (isElement(node) && node.namespaceURI === ns && node.localName === local) return node;
	return undefined;
}
const elementsOf = (parent: XmlElement): XmlElement[] =>
	Array.from(parent.childNodes).filter(isElement);

/** The first `wps:wsp` below a graphic's data. */
export function shapeElement(graphicData: XmlElement | undefined): XmlElement | undefined {
	return graphicData?.getElementsByTagNameNS(WPS_NS, 'wsp')[0] as XmlElement | undefined;
}

/** Text of each paragraph of a shape's `wps:txbx`, or undefined without one. */
export function shapeParagraphs(wsp: XmlElement | undefined): string[] | undefined {
	const content = child(child(wsp, WPS_NS, 'txbx'), W_NS, 'txbxContent');
	if (!content) return undefined;
	return Array.from(content.getElementsByTagNameNS(W_NS, 'p')).map((paragraph) =>
		Array.from(paragraph.getElementsByTagNameNS(W_NS, 't'))
			.map((t) => t.textContent ?? '')
			.join(''),
	);
}

/** Paragraphs of plain runs only (what the editor can write back without losing anything). */
function textIsPlain(wsp: XmlElement): boolean {
	const content = child(child(wsp, WPS_NS, 'txbx'), W_NS, 'txbxContent');
	if (!content) return true;
	const only = (node: XmlElement, allowed: string[]) =>
		elementsOf(node).every(
			(item) => item.namespaceURI === W_NS && allowed.includes(item.localName),
		);
	return (
		only(content, ['p']) &&
		Array.from(content.getElementsByTagNameNS(W_NS, 'p')).every(
			(paragraph) =>
				only(paragraph, ['pPr', 'r']) &&
				Array.from(paragraph.getElementsByTagNameNS(W_NS, 'r')).every((run) =>
					only(run, ['rPr', 't']),
				),
		)
	);
}

const LINE_CHILDREN = new Set([
	'solidFill',
	'noFill',
	'prstDash',
	'round',
	'miter',
	'bevel',
	'headEnd',
	'tailEnd',
]);
const FILLS = new Set(['solidFill', 'noFill']);

/** Whether `spPr` holds only what the shape model writes: transform, preset, a plain fill and outline. */
function spPrIsSimple(spPr: XmlElement): boolean {
	return elementsOf(spPr).every((item) => {
		if (item.namespaceURI !== A_NS) return false;
		if (item.localName === 'xfrm' || item.localName === 'prstGeom' || item.localName === 'custGeom')
			return true;
		if (FILLS.has(item.localName)) return true;
		if (item.localName !== 'ln') return false;
		return elementsOf(item).every(
			(part) => part.namespaceURI === A_NS && LINE_CHILDREN.has(part.localName),
		);
	});
}

function transform(spPr: XmlElement): Pick<ShapeSpec, 'rotation' | 'flipH' | 'flipV'> {
	const xfrm = child(spPr, A_NS, 'xfrm');
	const result: Pick<ShapeSpec, 'rotation' | 'flipH' | 'flipV'> = {};
	const rot = Number(xfrm?.getAttribute('rot'));
	if (Number.isFinite(rot) && rot !== 0) result.rotation = Math.round((rot / 60000) * 100) / 100;
	if (xfrm?.getAttribute('flipH') === '1') result.flipH = true;
	if (xfrm?.getAttribute('flipV') === '1') result.flipV = true;
	return result;
}

/**
 * The shape in a `wps:wsp` that is not a text box, or undefined when it has no preset geometry the
 * editor can draw. Editing a shape Word wrote with a VML fallback (`mc:AlternateContent`) replaces
 * the pair with the modern drawing alone.
 */
export function parseShape(graphicData: XmlElement | undefined): ShapeSpec | undefined {
	const wsp = shapeElement(graphicData);
	if (!wsp) return undefined;
	if (child(wsp, WPS_NS, 'cNvSpPr')?.getAttribute('txBox') === '1') return undefined;
	const spPr = child(wsp, WPS_NS, 'spPr');
	if (!spPr) return undefined;
	const custom = child(spPr, A_NS, 'custGeom');
	const geometry = parseDrawingGeometry(spPr);
	if (geometry.issue || (!custom && !lookupPresetShape(geometry.geometry))) return undefined;
	const fill = parseDrawingFill(spPr);
	const ln = child(spPr, A_NS, 'ln');
	const line: DrawingLine | undefined = ln ? parseDrawingLine(ln) : undefined;
	const text = shapeParagraphs(wsp);
	const simple = spPrIsSimple(spPr);
	return {
		preset: geometry.geometry,
		...(custom ? { customGeometryXml: buildXml(custom) } : {}),
		...(geometry.adjustments && Object.keys(geometry.adjustments).length
			? { adjustments: geometry.adjustments }
			: {}),
		...(fill && fill.kind !== 'unsupported' ? { fill } : {}),
		...(line ? { line } : {}),
		...(text ? { text } : {}),
		...transform(spPr),
		editable: simple,
		textEditable: simple && textIsPlain(wsp),
	};
}
