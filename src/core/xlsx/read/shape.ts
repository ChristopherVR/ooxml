// Reads the DrawingML shapes of a worksheet drawing (`xdr:sp`, `xdr:cxnSp`, `xdr:grpSp`) into the
// model the painter draws, through the shared drawingml readers. The XML itself is kept verbatim
// by the drawing reader, so nothing here is written back.
import { parseDrawingShapeStyle } from '../../drawingml/shape-style-refs';
import { parseDrawingFill, parseDrawingLine } from '../../drawingml/drawing-fill';
import { parseDrawingGeometry, parseTransform } from '../../drawingml/drawing-geometry';
import { parseDrawingTextBody } from '../../drawingml/drawing-text';
import { NS, elements, first, type XmlElement } from '../../xml/index';
import type { ShapeNode } from '../model-shapes';
import { readWordArtShape } from '../word-art';
import { equationShapeBranches, readEquationParagraphs } from '../shape-equations';
import { att, boolAttr } from './xml-util';
import { readDrawingShadow } from './drawing-format';

const xdr = (parent: XmlElement | undefined, local: string) => first(parent, local, NS.xdr);

/** Reads one shape, connector or group; children of a group are read recursively. */
export function readShapeNode(element: XmlElement): ShapeNode | undefined {
	const branches = equationShapeBranches(element);
	if (branches) element = branches.shape;
	const local = element.localName;
	if (local !== 'sp' && local !== 'cxnSp' && local !== 'grpSp') return undefined;
	const isGroup = local === 'grpSp';
	const nv = elements(element).find((node) => node.localName?.startsWith('nv'));
	const cNvPr = xdr(nv, 'cNvPr');
	const spPr = xdr(element, isGroup ? 'grpSpPr' : 'spPr');
	const xfrm = first(spPr, 'xfrm', NS.a);
	const transform = parseTransform(xfrm);
	const geometry = isGroup ? { geometry: 'group' } : parseDrawingGeometry(spPr);
	const node: ShapeNode = {
		kind: isGroup ? 'group' : local === 'cxnSp' ? 'connector' : 'shape',
		frame: transform?.frame ?? { x: 0, y: 0, width: 0, height: 0 },
		geometry: geometry.geometry,
	};
	const name = att(cNvPr, 'name');
	if (name) node.name = name;
	const description = att(cNvPr, 'descr');
	if (description) node.description = description;
	if (boolAttr(cNvPr, 'hidden')) node.hidden = true;
	if (att(xdr(nv, 'cNvSpPr'), 'txBox') === '1') node.textBox = true;
	if (transform?.rotation) node.rotation = transform.rotation;
	if (transform?.flipHorizontal) node.flipH = true;
	if (transform?.flipVertical) node.flipV = true;
	if ('adjustments' in geometry && geometry.adjustments) node.adjustments = geometry.adjustments;
	if ('paths' in geometry && geometry.paths) node.paths = geometry.paths;
	if (!isGroup) {
		const fill = parseDrawingFill(spPr);
		if (fill) node.fill = fill;
		const line = parseDrawingLine(first(spPr, 'ln', NS.a));
		if (line) node.line = line;
		const shadow = readDrawingShadow(spPr);
		if (shadow) node.shadow = shadow;
		if (first(spPr, 'effectDag', NS.a)) node.effectDag = true;
		if (node.effectDag || first(spPr, 'effectLst', NS.a)) node.hasExplicitEffects = true;
		const style = parseDrawingShapeStyle(xdr(element, 'style'));
		if (style) node.style = style;
		const body = xdr(element, 'txBody');
		const text = parseDrawingTextBody(body);
		const equations = readEquationParagraphs(element, branches?.fallback);
		if (equations) node.equationParagraphs = equations;
		if (text && (text.text || equations)) node.text = text;
		const wordArt = node.textBox && readWordArtShape(body);
		if (wordArt) node.wordArt = wordArt;
		return node;
	}
	const childFrame = parseFrameOf(first(xfrm, 'chOff', NS.a), first(xfrm, 'chExt', NS.a));
	if (childFrame) node.childFrame = childFrame;
	node.children = elements(element).flatMap((child) => {
		const read = readShapeNode(child);
		return read ? [read] : [];
	});
	return node;
}

function parseFrameOf(off: XmlElement | undefined, ext: XmlElement | undefined) {
	if (!off || !ext) return undefined;
	return {
		x: Number(att(off, 'x') ?? 0),
		y: Number(att(off, 'y') ?? 0),
		width: Number(att(ext, 'cx') ?? 0),
		height: Number(att(ext, 'cy') ?? 0),
	};
}
