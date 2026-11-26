// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// In-place patches of an existing drawing's XML when its picture, chart or run keeps its markup:
// the extent, the docPr alt text and title, a floating picture's position, and clearing a run
// down to its properties and the drawing it keeps.
import { writeDrawingTransform } from './drawing-transform';
import type { DrawingTransform } from './drawing-arrange';
import { WPG_NS } from './group-shape';
import { WPS_NS } from './shape-parse';
import { EMU_PER_PIXEL } from './drawing';
import type { PicturePlacement } from './model';
import { isElement, makeNS, named, type XmlDocument, type XmlElement } from './xml';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

function walk(root: XmlElement, visit: (node: XmlElement) => void): void {
	const stack: XmlElement[] = [root];
	while (stack.length) {
		const node = stack.shift()!;
		visit(node);
		for (const child of Array.from(node.childNodes)) if (isElement(child)) stack.push(child);
	}
}

export function updateExtent(drawing: XmlElement, widthPx: number, heightPx: number): void {
	const cx = String(Math.max(1, Math.round(widthPx * EMU_PER_PIXEL)));
	const cy = String(Math.max(1, Math.round(heightPx * EMU_PER_PIXEL)));
	walk(drawing, (node) => {
		const isExtent = node.namespaceURI === WP_NS && node.localName === 'extent';
		// Only a:xfrm/a:ext is a size; a:extLst/a:ext carries a uri and extension content.
		const parent = node.parentNode as XmlElement | null;
		const isExt =
			node.namespaceURI === A_NS &&
			node.localName === 'ext' &&
			parent?.namespaceURI === A_NS &&
			parent.localName === 'xfrm';
		if (isExtent || isExt) {
			node.setAttribute('cx', cx);
			node.setAttribute('cy', cy);
		}
	});
}

export function updateDocPr(
	drawing: XmlElement,
	altText: string | undefined,
	title: string | undefined,
	drawingName?: string,
): void {
	// The holder's docPr precedes the graphic and any drawings in its text box story.
	const node = drawing.getElementsByTagNameNS(WP_NS, 'docPr')[0];
	if (!node) return;
	if (altText) node.setAttribute('descr', altText);
	else node.removeAttribute('descr');
	if (title) node.setAttribute('title', title);
	else node.removeAttribute('title');
	if (drawingName) node.setAttribute('name', drawingName);
}

/** Rewrites one `wp:positionH` or `wp:positionV` from the model's alignment or offset. */
function patchPosition(
	doc: XmlDocument,
	anchor: XmlElement,
	name: 'positionH' | 'positionV',
	relativeFrom: string | undefined,
	align: string | undefined,
	offsetPx: number | undefined,
): void {
	const position = Array.from(anchor.childNodes).find(
		(child): child is XmlElement =>
			isElement(child) && child.namespaceURI === WP_NS && child.localName === name,
	);
	if (!position) return;
	if (relativeFrom) position.setAttribute('relativeFrom', relativeFrom);
	for (const child of Array.from(position.childNodes)) position.removeChild(child);
	const value = align ? makeNS(doc, WP_NS, 'wp:align') : makeNS(doc, WP_NS, 'wp:posOffset');
	value.appendChild(
		doc.createTextNode(align ?? String(Math.round((offsetPx ?? 0) * EMU_PER_PIXEL))),
	);
	position.appendChild(value);
}

/** Moves a floating picture: only its position elements change; wrapping and the rest stay. */
export function updatePlacement(
	drawing: XmlElement,
	doc: XmlDocument,
	placement: PicturePlacement,
	previous?: PicturePlacement,
): void {
	walk(drawing, (node) => {
		if (node.namespaceURI !== WP_NS || node.localName !== 'anchor' || node.parentNode !== drawing)
			return;
		node.setAttribute('simplePos', '0');
		if (
			!previous ||
			placement.relativeFrom !== previous.relativeFrom ||
			placement.align !== previous.align ||
			placement.offsetXPx !== previous.offsetXPx
		)
			patchPosition(
				doc,
				node,
				'positionH',
				placement.relativeFrom,
				placement.align,
				placement.offsetXPx,
			);
		if (
			!previous ||
			placement.relativeFromV !== previous.relativeFromV ||
			placement.alignV !== previous.alignV ||
			placement.offsetYPx !== previous.offsetYPx
		)
			patchPosition(
				doc,
				node,
				'positionV',
				placement.relativeFromV,
				placement.alignV,
				placement.offsetYPx,
			);
	});
}

/** Patches only the outer shape or picture, retaining any transforms inside a group. */
export function updateDrawingTransform(drawing: XmlElement, transform: DrawingTransform): void {
	const graphic = drawing.getElementsByTagNameNS(A_NS, 'graphicData')[0];
	const ns = graphic?.getAttribute('uri');
	if (!graphic || !ns) return;
	const properties = graphic.getElementsByTagNameNS(ns, ns === WPG_NS ? 'grpSpPr' : 'spPr')[0];
	const xfrm = properties?.getElementsByTagNameNS(A_NS, 'xfrm')[0];
	if (xfrm) writeDrawingTransform(xfrm, transform);
}

export function updateTextBoxDirection(drawing: XmlElement, vertical: string | undefined): void {
	drawing.getElementsByTagNameNS(WPS_NS, 'bodyPr')[0]?.setAttribute('vert', vertical ?? 'horz');
}

export function clearNonProperties(node: XmlElement, keep?: XmlElement): void {
	for (const child of Array.from(node.childNodes)) {
		if (isElement(child) && (named(child, 'rPr') || child === keep)) continue;
		node.removeChild(child);
	}
}
