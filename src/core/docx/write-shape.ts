// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes drawing shapes (`wps:wsp`): a new inline or floating shape, and an edited one rebuilt from
// its model while keeping the identity (`wp:docPr`), style, text-body properties and any text the
// model cannot edit from the drawing it replaces.
import { groupElement, shapeElement } from './write-group-shape';
import { WPG_NS } from './group-shape';
import type { InlineImage } from './model';
import { A_NS, WPS_NS } from './shape-parse';
import { EMU_PER_PIXEL } from './units';
import { positionElement as position, wrapElement } from './write-drawing-wrap';
import { makeNS, makeW, type XmlDocument, type XmlElement } from './xml';

const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const el = (doc: XmlDocument, ns: string, name: string) => makeNS(doc, ns, name);
const add = (parent: XmlElement, ...items: XmlElement[]) => {
	for (const item of items) parent.appendChild(item);
	return parent;
};
const emu = (px: number) => String(Math.max(1, Math.round(px * EMU_PER_PIXEL)));
const attrs = (element: XmlElement, values: Record<string, string>) => {
	for (const [name, value] of Object.entries(values)) element.setAttribute(name, value);
	return element;
};

export interface ShapeIdentity {
	id: string;
	name?: string;
	descr?: string;
	title?: string;
}

/** A complete `w:drawing` for the shape, floating when the image is anchored. */
export function buildShapeDrawing(
	doc: XmlDocument,
	image: InlineImage,
	identity: ShapeIdentity,
	relativeHeight = 251658240,
): XmlElement {
	const spec = image.shape!;
	const drawing = makeW(doc, 'drawing');
	const placement = image.placement;
	const anchored = Boolean(image.anchored);
	const holder = el(doc, WP_NS, anchored ? 'wp:anchor' : 'wp:inline');
	if (anchored) {
		attrs(holder, {
			distT: '0',
			distB: '0',
			distL: '114300',
			distR: '114300',
			simplePos: '0',
			relativeHeight: String(placement?.relativeHeight ?? relativeHeight),
			behindDoc: placement?.behindText ? '1' : '0',
			locked: '0',
			layoutInCell: placement?.layoutInCell === false ? '0' : '1',
			allowOverlap: '1',
		});
		add(
			holder,
			attrs(el(doc, WP_NS, 'wp:simplePos'), { x: '0', y: '0' }),
			position(
				doc,
				'H',
				placement?.relativeFrom ?? 'column',
				placement?.align,
				placement?.offsetXPx,
			),
			position(
				doc,
				'V',
				placement?.relativeFromV ?? 'paragraph',
				placement?.alignV,
				placement?.offsetYPx,
			),
		);
	} else for (const side of ['distT', 'distB', 'distL', 'distR']) holder.setAttribute(side, '0');
	add(
		holder,
		attrs(el(doc, WP_NS, 'wp:extent'), { cx: emu(image.widthPx), cy: emu(image.heightPx) }),
		attrs(el(doc, WP_NS, 'wp:effectExtent'), { l: '0', t: '0', r: '0', b: '0' }),
	);
	if (anchored) add(holder, wrapElement(doc, placement ?? { wrap: 'none' }));
	const docPr = attrs(el(doc, WP_NS, 'wp:docPr'), {
		id: identity.id,
		name: image.drawingName ?? identity.name ?? `Shape ${identity.id}`,
	});
	if (identity.descr) docPr.setAttribute('descr', identity.descr);
	if (identity.title) docPr.setAttribute('title', identity.title);
	const graphicData = attrs(el(doc, A_NS, 'a:graphicData'), {
		uri: spec.groupChildren ? WPG_NS : WPS_NS,
	});
	add(
		graphicData,
		spec.groupChildren ? groupElement(doc, image, spec) : shapeElement(doc, image, spec),
	);
	add(
		holder,
		docPr,
		el(doc, WP_NS, 'wp:cNvGraphicFramePr'),
		add(el(doc, A_NS, 'a:graphic'), graphicData),
	);
	return add(drawing, holder);
}

const firstNs = (root: XmlElement, ns: string, local: string) =>
	root.getElementsByTagNameNS(ns, local)[0] as XmlElement | undefined;

/**
 * A new drawing for an edited shape. The shape's identity, its `wps:style` and `wps:bodyPr` come from
 * `old`, and so does its text when the model cannot edit it.
 */
export function rebuildShapeDrawing(
	doc: XmlDocument,
	image: InlineImage,
	old: XmlElement,
): XmlElement {
	const oldDocPr = firstNs(old, WP_NS, 'docPr');
	const identity: ShapeIdentity = {
		id: oldDocPr?.getAttribute('id') ?? '1',
		...(oldDocPr?.getAttribute('name') ? { name: oldDocPr.getAttribute('name')! } : {}),
		...(oldDocPr?.getAttribute('descr') ? { descr: oldDocPr.getAttribute('descr')! } : {}),
		...(oldDocPr?.getAttribute('title') ? { title: oldDocPr.getAttribute('title')! } : {}),
	};
	const oldAnchor = firstNs(old, WP_NS, 'anchor');
	const relativeHeight = Number(oldAnchor?.getAttribute('relativeHeight')) || undefined;
	const next = buildShapeDrawing(doc, image, identity, relativeHeight);
	const wsp = firstNs(next, WPS_NS, 'wsp')!;
	const keep = (local: string) => {
		const previous = firstNs(old, WPS_NS, local);
		const current = firstNs(next, WPS_NS, local);
		if (previous && current) wsp.replaceChild(previous.cloneNode(true), current);
	};
	keep('style');
	keep('bodyPr');
	if (!image.shape?.textEditable) {
		const text = firstNs(old, WPS_NS, 'txbx');
		const current = firstNs(next, WPS_NS, 'txbx');
		if (text && current) wsp.replaceChild(text.cloneNode(true), current);
		else if (text) wsp.insertBefore(text.cloneNode(true), firstNs(next, WPS_NS, 'bodyPr') ?? null);
	}
	return next;
}
