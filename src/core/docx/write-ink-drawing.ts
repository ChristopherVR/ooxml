// The drawing of Word ink, as Word 16 writes it: `w:drawing` around a `wp:inline` (or, floating,
// a `wp:anchor`) whose `a:graphicData` (uri `wordprocessingInk`) holds a `w14:contentPart` naming
// the InkML part, with its box in `w14:xfrm`. Word opens it without an `mc:AlternateContent`
// wrapper or a picture fallback, and writes it the same way when it saves.
import { EMU_PER_PIXEL } from './units';
import { W14_NS, WORD_INK_GRAPHIC_URI } from './ink';
import type { InlineImage } from './model';
import { applyDrawingWrap } from './write-drawing-wrap';
import { isElement, makeNS, makeW, REL_NS, type XmlDocument, type XmlElement } from './xml';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';

function child(parent: XmlElement, ns: string, name: string, attrs: Record<string, string> = {}) {
	const node = makeNS(parent.ownerDocument!, ns, name);
	for (const [key, value] of Object.entries(attrs)) node.setAttribute(key, value);
	parent.appendChild(node);
	return node;
}

const emu = (px: number): string => String(Math.max(1, Math.round(px * EMU_PER_PIXEL)));

/** A new ink drawing for the part related by `relId`, floating when the image is anchored. */
export function buildInkDrawing(
	doc: XmlDocument,
	image: InlineImage,
	relId: string,
	docPrId: string,
): XmlElement {
	const drawing = makeW(doc, 'drawing');
	const inline = child(drawing, WP_NS, 'wp:inline', {
		distT: '0',
		distB: '0',
		distL: '0',
		distR: '0',
	});
	const cx = emu(image.widthPx);
	const cy = emu(image.heightPx);
	child(inline, WP_NS, 'wp:extent', { cx, cy });
	child(inline, WP_NS, 'wp:effectExtent', { l: '0', t: '0', r: '0', b: '0' });
	const docPr = child(inline, WP_NS, 'wp:docPr', {
		id: docPrId,
		name: image.drawingName ?? `Ink ${docPrId}`,
	});
	if (image.altText) docPr.setAttribute('descr', image.altText);
	if (image.title) docPr.setAttribute('title', image.title);
	child(inline, WP_NS, 'wp:cNvGraphicFramePr');
	const graphic = child(inline, A_NS, 'a:graphic');
	const data = child(graphic, A_NS, 'a:graphicData', { uri: WORD_INK_GRAPHIC_URI });
	const part = child(data, W14_NS, 'w14:contentPart', { bwMode: 'auto' });
	part.setAttributeNS(REL_NS, 'r:id', relId);
	const nv = child(part, W14_NS, 'w14:nvContentPartPr');
	child(nv, W14_NS, 'w14:cNvContentPartPr');
	const xfrm = child(part, W14_NS, 'w14:xfrm');
	child(xfrm, A_NS, 'a:off', { x: '0', y: '0' });
	child(xfrm, A_NS, 'a:ext', { cx, cy });
	if (image.anchored)
		applyDrawingWrap(
			drawing,
			doc,
			true,
			image.placement ?? { wrap: 'none' },
			image.placement?.relativeHeight ?? 251658240 + (Number(docPrId) % 4096) * 1024,
		);
	return drawing;
}

/** Sets the size `w14:xfrm` gives the ink (the extent itself is patched with the drawing's). */
export function patchInkXfrm(drawing: XmlElement, widthPx: number, heightPx: number): void {
	for (const xfrm of Array.from(drawing.getElementsByTagNameNS(W14_NS, 'xfrm')))
		for (const node of Array.from(xfrm.childNodes))
			if (isElement(node) && node.namespaceURI === A_NS && node.localName === 'ext') {
				node.setAttribute('cx', emu(widthPx));
				node.setAttribute('cy', emu(heightPx));
			}
}
