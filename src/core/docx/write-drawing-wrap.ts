// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Wrap Text in the XML: the wrap element of a floating drawing, and turning a kept `wp:inline`
// into a `wp:anchor` (or back) while its extent, ids, `wp:docPr`, locks and graphic survive.
import type { PicturePlacement } from './inline-model';
import { EMU_PER_PIXEL } from './units';
import { isElement, makeNS, type XmlDocument, type XmlElement } from './xml';

const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const WRAP_NAMES = ['wrapNone', 'wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom'];
/** CT_Anchor children before the wrap choice; everything else follows it. */
const BEFORE_WRAP = ['simplePos', 'positionH', 'positionV', 'extent', 'effectExtent'];
/** CT_Anchor children with no place in CT_Inline. */
const ANCHOR_ONLY = ['simplePos', 'positionH', 'positionV', ...WRAP_NAMES];
/** Word's default text distance beside a floating picture: 0.125 inch. */
const SIDE_DISTANCE = '114300';

const el = (doc: XmlDocument, name: string) => makeNS(doc, WP_NS, name);
const attrs = (element: XmlElement, values: Record<string, string>) => {
	for (const [name, value] of Object.entries(values)) element.setAttribute(name, value);
	return element;
};
const wpChildren = (parent: XmlElement) =>
	Array.from(parent.childNodes).filter(
		(node): node is XmlElement => isElement(node) && node.namespaceURI === WP_NS,
	);

/** A rectangle wrap polygon in Word's 21600-unit square, marked unedited. */
function rectanglePolygon(doc: XmlDocument): XmlElement {
	const polygon = attrs(el(doc, 'wp:wrapPolygon'), { edited: '0' });
	const points = [
		['start', 0, 0],
		['lineTo', 0, 21600],
		['lineTo', 21600, 21600],
		['lineTo', 21600, 0],
		['lineTo', 0, 0],
	] as const;
	for (const [name, x, y] of points)
		polygon.appendChild(attrs(el(doc, `wp:${name}`), { x: String(x), y: String(y) }));
	return polygon;
}

/** The wrap element for a placement; tight and through reuse `polygon` when given one. */
export function wrapElement(
	doc: XmlDocument,
	placement: PicturePlacement,
	polygon?: XmlElement,
): XmlElement {
	const wrapText = placement.wrapText ?? 'bothSides';
	switch (placement.wrap) {
		case 'square':
			return attrs(el(doc, 'wp:wrapSquare'), { wrapText });
		case 'tight':
		case 'through': {
			const wrap = el(doc, placement.wrap === 'tight' ? 'wp:wrapTight' : 'wp:wrapThrough');
			wrap.setAttribute('wrapText', wrapText);
			wrap.appendChild(polygon ?? rectanglePolygon(doc));
			return wrap;
		}
		case 'topAndBottom':
			return el(doc, 'wp:wrapTopAndBottom');
		default:
			return el(doc, 'wp:wrapNone');
	}
}

/** A `wp:positionH` or `wp:positionV` from an alignment or an offset in CSS pixels. */
export function positionElement(
	doc: XmlDocument,
	axis: 'H' | 'V',
	relativeFrom: string,
	align: string | undefined,
	offsetPx: number | undefined,
): XmlElement {
	const node = attrs(el(doc, `wp:position${axis}`), { relativeFrom });
	const value = el(doc, align ? 'wp:align' : 'wp:posOffset');
	value.appendChild(
		doc.createTextNode(align ?? String(Math.round((offsetPx ?? 0) * EMU_PER_PIXEL))),
	);
	node.appendChild(value);
	return node;
}

/** Swaps the anchor's wrap element for `placement`'s and sets `behindDoc`. */
function patchAnchorWrap(doc: XmlDocument, anchor: XmlElement, placement: PicturePlacement): void {
	const children = wpChildren(anchor);
	const old = children.find((child) => WRAP_NAMES.includes(child.localName));
	const oldPolygon = old && wpChildren(old).find((child) => child.localName === 'wrapPolygon');
	const next = wrapElement(doc, placement, oldPolygon);
	if (old) anchor.replaceChild(next, old);
	else {
		const after = children.find((child) => !BEFORE_WRAP.includes(child.localName));
		anchor.insertBefore(next, after ?? null);
	}
	anchor.setAttribute('behindDoc', placement.wrap === 'none' && placement.behindText ? '1' : '0');
}

/** Moves every child and the non-anchor attributes of `from` onto `to`, then swaps them. */
function replaceHolder(from: XmlElement, to: XmlElement, skip: readonly string[]): void {
	for (const attribute of Array.from(from.attributes))
		if (!skip.includes(attribute.name) && !to.hasAttribute(attribute.name))
			to.setAttributeNS(attribute.namespaceURI, attribute.name, attribute.value);
	for (const child of Array.from(from.childNodes)) to.appendChild(child);
	from.parentNode?.replaceChild(to, from);
}

const ANCHOR_ATTRIBUTES = [
	'distT',
	'distB',
	'distL',
	'distR',
	'simplePos',
	'relativeHeight',
	'behindDoc',
	'locked',
	'layoutInCell',
	'allowOverlap',
	'hidden',
];

/** `wp:inline` to `wp:anchor`, with Word's floating defaults and the placement's position. */
function inlineToAnchor(
	doc: XmlDocument,
	inline: XmlElement,
	placement: PicturePlacement,
	relativeHeight: number,
): void {
	const anchor = attrs(el(doc, 'wp:anchor'), {
		distT: inline.getAttribute('distT') ?? '0',
		distB: inline.getAttribute('distB') ?? '0',
		distL: SIDE_DISTANCE,
		distR: SIDE_DISTANCE,
		simplePos: '0',
		relativeHeight: String(relativeHeight),
		behindDoc: '0',
		locked: '0',
		layoutInCell: placement.layoutInCell === false ? '0' : '1',
		allowOverlap: '1',
	});
	anchor.appendChild(attrs(el(doc, 'wp:simplePos'), { x: '0', y: '0' }));
	anchor.appendChild(
		positionElement(
			doc,
			'H',
			placement.relativeFrom ?? 'column',
			placement.align,
			placement.offsetXPx,
		),
	);
	anchor.appendChild(
		positionElement(
			doc,
			'V',
			placement.relativeFromV ?? 'paragraph',
			placement.alignV,
			placement.offsetYPx,
		),
	);
	replaceHolder(inline, anchor, ANCHOR_ATTRIBUTES);
	patchAnchorWrap(doc, anchor, placement);
}

/** `wp:anchor` to `wp:inline`: positions, wrapping and relative sizes go; the rest stays. */
function anchorToInline(doc: XmlDocument, anchor: XmlElement): void {
	const inline = attrs(el(doc, 'wp:inline'), { distT: '0', distB: '0', distL: '0', distR: '0' });
	for (const child of wpChildren(anchor))
		if (ANCHOR_ONLY.includes(child.localName)) anchor.removeChild(child);
	// Word 2010's relative sizes (`wp14:sizeRelH`/`V`) only exist on a floating drawing.
	for (const child of Array.from(anchor.childNodes))
		if (isElement(child) && child.localName.startsWith('sizeRel')) anchor.removeChild(child);
	replaceHolder(anchor, inline, ANCHOR_ATTRIBUTES);
}

/**
 * Applies a Wrap Text change to a kept drawing: In Line with Text writes `wp:inline`, every other
 * choice `wp:anchor` with the placement's wrap element. Returns false when the drawing has neither
 * (a legacy VML picture), so the caller can keep it as it was.
 */
export function applyDrawingWrap(
	drawing: XmlElement,
	doc: XmlDocument,
	anchored: boolean,
	placement: PicturePlacement | undefined,
	relativeHeight = 251658240,
): boolean {
	const holder = wpChildren(drawing).find(
		(child) => child.localName === 'inline' || child.localName === 'anchor',
	);
	if (!holder) return false;
	if (!anchored) {
		if (holder.localName === 'anchor') anchorToInline(doc, holder);
		return true;
	}
	const floating = placement ?? { wrap: 'none' as const };
	if (holder.localName === 'inline') inlineToAnchor(doc, holder, floating, relativeHeight);
	else patchAnchorWrap(doc, holder, floating);
	return true;
}

const wrapKey = (anchored: boolean | undefined, placement: PicturePlacement | undefined) =>
	anchored
		? `${placement?.wrap ?? 'none'}|${placement?.behindText ? 1 : 0}|${placement?.wrapText ?? ''}`
		: 'inline';

/** A z-order for a newly floating drawing, kept apart from its neighbours by its `wp:docPr` id. */
function newRelativeHeight(drawing: XmlElement): number {
	const docPr = drawing.getElementsByTagNameNS(WP_NS, 'docPr')[0];
	return 251658240 + ((Number(docPr?.getAttribute('id')) || 0) % 4096) * 1024;
}

interface WrapState {
	anchored?: boolean | undefined;
	placement?: PicturePlacement | undefined;
}

/** Writes a picture's Wrap Text change onto its drawing when the choice differs from `base`'s. */
export function patchDrawingWrap(
	drawing: XmlElement,
	doc: XmlDocument,
	image: WrapState,
	base?: WrapState,
): void {
	const next = wrapKey(image.anchored, image.placement);
	if (next === (base ? wrapKey(base.anchored, base.placement) : 'inline')) return;
	applyDrawingWrap(
		drawing,
		doc,
		Boolean(image.anchored),
		image.placement,
		image.placement?.relativeHeight ?? newRelativeHeight(drawing),
	);
}

/** Edits the top-level anchor's relative stacking order. */
export function patchDrawingOrder(
	drawing: XmlElement,
	next: PicturePlacement | undefined,
	base?: PicturePlacement,
): void {
	if (next?.relativeHeight === undefined || next.relativeHeight === base?.relativeHeight) return;
	wpChildren(drawing)
		.find((child) => child.localName === 'anchor')
		?.setAttribute('relativeHeight', String(next.relativeHeight));
}

const POSITION_KEYS = [
	'relativeFrom',
	'align',
	'offsetXPx',
	'relativeFromV',
	'alignV',
	'offsetYPx',
] as const;

/** Whether a floating picture moved: a wrap change alone leaves its position XML as it was. */
export function placementPositionChanged(
	next: PicturePlacement,
	base: PicturePlacement | undefined,
): boolean {
	return !base || POSITION_KEYS.some((key) => next[key] !== base[key]);
}
