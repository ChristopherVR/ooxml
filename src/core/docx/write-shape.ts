// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes drawing shapes (`wps:wsp`): a new inline or floating shape, and an edited one rebuilt from
// its model while keeping the identity (`wp:docPr`), style, text-body properties and any text the
// model cannot edit from the drawing it replaces.
import { drawingFillXml } from '../drawingml/index';
import type { DrawingFill } from '../drawingml/index';
import { drawingLineXml } from '../drawingml/write-line';
import { parseXml } from '../xml/index';
import type { InlineImage, PicturePlacement } from './model';
import type { ShapeSpec } from './shape-model';
import { A_NS, W_NS, WPS_NS } from './shape-parse';
import { EMU_PER_PIXEL } from './units';
import { makeNS, makeW, type XmlDocument, type XmlElement } from './xml';

const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const el = (doc: XmlDocument, ns: string, name: string) => makeNS(doc, ns, name);
const add = (parent: XmlElement, ...items: XmlElement[]) => {
	for (const item of items) parent.appendChild(item);
	return parent;
};
const emu = (px: number) => String(Math.max(1, Math.round(px * EMU_PER_PIXEL)));
const emuOffset = (px: number) => String(Math.round(px * EMU_PER_PIXEL));
const attrs = (element: XmlElement, values: Record<string, string>) => {
	for (const [name, value] of Object.entries(values)) element.setAttribute(name, value);
	return element;
};

/** The shared writers repeat `xmlns:a` on every fragment; the document's own declaration serves. */
function dropDeclarations(element: XmlElement): void {
	if (element.getAttribute('xmlns:a') === A_NS) element.removeAttribute('xmlns:a');
	for (const node of Array.from(element.childNodes))
		if (node.nodeType === 1) dropDeclarations(node as XmlElement);
}

function importFragment(doc: XmlDocument, xml: string): XmlElement[] {
	const wrapper = parseXml(`<w xmlns:a="${A_NS}">${xml}</w>`).documentElement;
	return Array.from(wrapper.childNodes)
		.filter((node): node is XmlElement => node.nodeType === 1)
		.map((node) => {
			const imported = doc.importNode(node, true) as XmlElement;
			dropDeclarations(imported);
			return imported;
		});
}

const srgb = (fill: DrawingFill | undefined): string | undefined =>
	fill?.kind === 'solid' && fill.color.kind === 'srgb' ? fill.color.value : undefined;

/** Text colour that stays readable on an explicit fill; the default style's white needs none. */
function explicitTextColor(spec: ShapeSpec): string | undefined {
	if (!spec.fill) return undefined;
	if (spec.fill.kind === 'none') return '000000';
	const hex = srgb(spec.fill);
	if (!hex || !/^[0-9a-f]{6}$/i.test(hex)) return undefined;
	const channel = (i: number) => Number.parseInt(hex.slice(i, i + 2), 16) / 255;
	const light = 0.2126 * channel(0) + 0.7152 * channel(2) + 0.0722 * channel(4) > 0.6;
	return light ? '000000' : undefined;
}

function textParagraphs(doc: XmlDocument, lines: readonly string[], color?: string): XmlElement[] {
	return lines.map((line) => {
		const paragraph = makeW(doc, 'p');
		const props = makeW(doc, 'pPr');
		const jc = makeW(doc, 'jc');
		jc.setAttributeNS(W_NS, 'w:val', 'center');
		add(paragraph, add(props, jc));
		if (line) {
			const run = makeW(doc, 'r');
			if (color) {
				const rPr = makeW(doc, 'rPr');
				const c = makeW(doc, 'color');
				c.setAttributeNS(W_NS, 'w:val', color);
				add(run, add(rPr, c));
			}
			const t = makeW(doc, 't');
			if (/^\s|\s$/.test(line))
				t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
			t.textContent = line;
			add(paragraph, add(run, t));
		}
		return paragraph;
	});
}

function shapeStyle(doc: XmlDocument): XmlElement {
	const ref = (name: string, idx: string, scheme: string, shade?: string) => {
		const node = attrs(el(doc, A_NS, `a:${name}`), { idx });
		const color = attrs(el(doc, A_NS, 'a:schemeClr'), { val: scheme });
		if (shade) add(color, attrs(el(doc, A_NS, 'a:shade'), { val: shade }));
		return add(node, color);
	};
	return add(
		el(doc, WPS_NS, 'wps:style'),
		ref('lnRef', '2', 'accent1', '50000'),
		ref('fillRef', '1', 'accent1'),
		ref('effectRef', '0', 'accent1'),
		ref('fontRef', 'minor', 'lt1'),
	);
}

function wrapElement(doc: XmlDocument, placement: PicturePlacement): XmlElement {
	switch (placement.wrap) {
		case 'square':
			return attrs(el(doc, WP_NS, 'wp:wrapSquare'), { wrapText: 'bothSides' });
		case 'tight':
		case 'through': {
			const wrap = attrs(
				el(doc, WP_NS, `wp:wrap${placement.wrap === 'tight' ? 'Tight' : 'Through'}`),
				{
					wrapText: 'bothSides',
				},
			);
			const polygon = attrs(el(doc, WP_NS, 'wp:wrapPolygon'), { edited: '0' });
			const point = (name: string, x: string, y: string) =>
				attrs(el(doc, WP_NS, `wp:${name}`), { x, y });
			return add(
				wrap,
				add(
					polygon,
					point('start', '0', '0'),
					point('lineTo', '0', '21600'),
					point('lineTo', '21600', '21600'),
					point('lineTo', '21600', '0'),
					point('lineTo', '0', '0'),
				),
			);
		}
		case 'topAndBottom':
			return el(doc, WP_NS, 'wp:wrapTopAndBottom');
		default:
			return el(doc, WP_NS, 'wp:wrapNone');
	}
}

function position(
	doc: XmlDocument,
	axis: 'H' | 'V',
	relativeFrom: string,
	align: string | undefined,
	offsetPx: number | undefined,
): XmlElement {
	const node = attrs(el(doc, WP_NS, `wp:position${axis}`), { relativeFrom });
	if (align) {
		const a = el(doc, WP_NS, 'wp:align');
		a.textContent = align;
		return add(node, a);
	}
	const offset = el(doc, WP_NS, 'wp:posOffset');
	offset.textContent = emuOffset(offsetPx ?? 0);
	return add(node, offset);
}

function geometry(doc: XmlDocument, spec: ShapeSpec): XmlElement[] {
	const prst = attrs(el(doc, A_NS, 'a:prstGeom'), { prst: spec.preset });
	const avLst = el(doc, A_NS, 'a:avLst');
	for (const [name, value] of Object.entries(spec.adjustments ?? {}))
		add(avLst, attrs(el(doc, A_NS, 'a:gd'), { name, fmla: `val ${Math.round(value)}` }));
	const nodes: XmlElement[] = [add(prst, avLst)];
	if (spec.fill) {
		const xml = drawingFillXml(spec.fill);
		if (xml) nodes.push(...importFragment(doc, xml));
	}
	if (spec.line) nodes.push(...importFragment(doc, drawingLineXml(spec.line)));
	return nodes;
}

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
			relativeHeight: String(relativeHeight),
			behindDoc: placement?.behindText ? '1' : '0',
			locked: '0',
			layoutInCell: '1',
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
		name: identity.name ?? `Shape ${identity.id}`,
	});
	if (identity.descr) docPr.setAttribute('descr', identity.descr);
	if (identity.title) docPr.setAttribute('title', identity.title);
	const xfrm = attrs(el(doc, A_NS, 'a:xfrm'), {});
	if (spec.rotation) xfrm.setAttribute('rot', String(Math.round(spec.rotation * 60000)));
	if (spec.flipH) xfrm.setAttribute('flipH', '1');
	if (spec.flipV) xfrm.setAttribute('flipV', '1');
	add(
		xfrm,
		attrs(el(doc, A_NS, 'a:off'), { x: '0', y: '0' }),
		attrs(el(doc, A_NS, 'a:ext'), { cx: emu(image.widthPx), cy: emu(image.heightPx) }),
	);
	const spPr = add(el(doc, WPS_NS, 'wps:spPr'), xfrm, ...geometry(doc, spec));
	const wsp = add(
		el(doc, WPS_NS, 'wps:wsp'),
		el(doc, WPS_NS, 'wps:cNvSpPr'),
		spPr,
		shapeStyle(doc),
	);
	if (spec.text?.length) {
		const content = add(
			el(doc, W_NS, 'w:txbxContent'),
			...textParagraphs(doc, spec.text, explicitTextColor(spec)),
		);
		add(wsp, add(el(doc, WPS_NS, 'wps:txbx'), content));
	}
	const body = attrs(el(doc, WPS_NS, 'wps:bodyPr'), {
		rot: '0',
		vert: 'horz',
		wrap: 'square',
		lIns: '91440',
		tIns: '45720',
		rIns: '91440',
		bIns: '45720',
		anchor: 'ctr',
		anchorCtr: '0',
	});
	add(wsp, add(body, el(doc, A_NS, 'a:noAutofit')));
	const graphicData = attrs(el(doc, A_NS, 'a:graphicData'), { uri: WPS_NS });
	add(
		holder,
		docPr,
		el(doc, WP_NS, 'wp:cNvGraphicFramePr'),
		add(el(doc, A_NS, 'a:graphic'), add(graphicData, wsp)),
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
