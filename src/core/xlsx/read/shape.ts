// Reads the DrawingML shapes of a worksheet drawing (`xdr:sp`, `xdr:cxnSp`, `xdr:grpSp`) into the
// model the painter draws, through the shared drawingml readers. The XML itself is kept verbatim
// by the drawing reader, so nothing here is written back.
import { parseDrawingColorIn } from '../../drawingml/drawing-color';
import { parseDrawingFill, parseDrawingLine } from '../../drawingml/drawing-fill';
import { parseDrawingGeometry, parseTransform } from '../../drawingml/drawing-geometry';
import { parseDrawingTextBody } from '../../drawingml/drawing-text';
import { NS, elements, first, type XmlElement } from '../../xml/index';
import type { ShapeNode } from '../model-shapes';
import { att } from './xml-util';

const xdr = (parent: XmlElement | undefined, local: string) => first(parent, local, NS.xdr);

function styleRef(style: XmlElement | undefined, name: string) {
	const ref = first(style, name, NS.a);
	if (!ref) return undefined;
	const idx = Number(att(ref, 'idx') ?? 0);
	const color = parseDrawingColorIn(ref);
	return { idx: Number.isFinite(idx) ? idx : 0, ...(color ? { color } : {}) };
}

function readStyle(style: XmlElement | undefined): ShapeNode['style'] | undefined {
	if (!style) return undefined;
	const fill = styleRef(style, 'fillRef');
	const line = styleRef(style, 'lnRef');
	const fontColor = parseDrawingColorIn(first(style, 'fontRef', NS.a));
	return {
		...(fill ? { fill } : {}),
		...(line ? { line } : {}),
		...(fontColor ? { font: { color: fontColor } } : {}),
	};
}

/** Reads one shape, connector or group; children of a group are read recursively. */
export function readShapeNode(element: XmlElement): ShapeNode | undefined {
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
		const style = readStyle(xdr(element, 'style'));
		if (style) node.style = style;
		const text = parseDrawingTextBody(xdr(element, 'txBody'));
		if (text?.text) node.text = text;
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
