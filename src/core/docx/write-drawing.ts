import { applyPictureCrop } from '../drawingml/picture-crop';
import { applyPictureEffects } from '../drawingml/picture-effects';
import { applyPictureLayout } from '../drawingml/picture-layout';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes inline picture runs: reuses original wp:inline/pic:pic XML verbatim (only patching
// extent/alt text) for unchanged pictures, and synthesizes a minimal valid drawing for insertions.
import type { InlineImage } from './model';
import { EMU_PER_PIXEL } from './drawing';
import { buildWatermarkRun } from './watermark';
import { buildTextBoxDrawing, patchTextBox } from './write-text-box';
import { buildShapeDrawing, rebuildShapeDrawing } from './write-shape';
import { chartRun } from './write-chart-drawing';
import {
	clearNonProperties,
	updateDocPr,
	updateExtent,
	updatePlacement,
} from './write-drawing-patch';
import type { RelationshipAllocator } from './relationship-allocator';
import { patchFallbackSize } from './write-fallback';
import { diagramRun } from './write-diagram-drawing';
import { isElement, makeNS, makeW, named, REL_NS, type XmlDocument, type XmlElement } from './xml';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
const PICTURE_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

/** The drawing in an `mc:AlternateContent` run child: the `mc:Choice` Word 2010+ reads. */
export function alternateContentDrawing(content: XmlElement): XmlElement | undefined {
	const choice = Array.from(content.childNodes)
		.filter(isElement)
		.find((node) => node.localName === 'Choice');
	return choice && findDrawingChild(choice);
}

/** The run child that holds a drawing: the drawing itself, or the `mc:AlternateContent` around it. */
function drawingContainer(run: XmlElement | undefined): XmlElement | undefined {
	if (!run) return undefined;
	for (const child of Array.from(run.childNodes))
		if (isElement(child) && (named(child, 'drawing') || named(child, 'pict'))) return child;
	for (const child of Array.from(run.childNodes))
		if (
			isElement(child) &&
			child.localName === 'AlternateContent' &&
			alternateContentDrawing(child)
		)
			return child;
	return undefined;
}

export function findDrawingChild(run: XmlElement | undefined): XmlElement | undefined {
	if (!run) return undefined;
	for (const child of Array.from(run.childNodes)) {
		if (!isElement(child)) continue;
		if (named(child, 'drawing') || named(child, 'pict')) return child;
		if (child.localName === 'AlternateContent') {
			const inner = alternateContentDrawing(child);
			if (inner) return inner;
		}
	}
	return undefined;
}

function sameUnderlyingImage(a: InlineImage, b: InlineImage): boolean {
	return Boolean(a.partName) && a.relId === b.relId && a.partName === b.partName;
}

function el(doc: XmlDocument, ns: string, qualified: string): XmlElement {
	return makeNS(doc, ns, qualified);
}
/** @xmldom/xmldom elements only implement `appendChild`, not the newer `append(...)` convenience. */
function appendAll(parent: XmlElement, ...items: XmlElement[]): void {
	for (const item of items) parent.appendChild(item);
}

const SVG_EXTENSION_URI = '{96DAC541-7B7A-43D3-8B79-37D633B846F1}';
const SVG_NS = 'http://schemas.microsoft.com/office/drawing/2016/SVG/main';

/** Word 2016+ SVG pictures: the SVG rides in an a:blip extension beside its PNG fallback. */
function svgExtension(doc: XmlDocument, svgRelId: string): XmlElement {
	const extLst = el(doc, A_NS, 'a:extLst');
	const ext = el(doc, A_NS, 'a:ext');
	ext.setAttribute('uri', SVG_EXTENSION_URI);
	const svgBlip = doc.createElementNS(SVG_NS, 'asvg:svgBlip');
	svgBlip.setAttributeNS(REL_NS, 'r:embed', svgRelId);
	ext.appendChild(svgBlip);
	extLst.appendChild(ext);
	return extLst;
}

function buildInlineDrawing(
	doc: XmlDocument,
	image: InlineImage,
	relId: string,
	docPrId: string,
	svgRelId?: string,
): XmlElement {
	const drawing = makeW(doc, 'drawing');
	const inline = el(doc, WP_NS, 'wp:inline');
	for (const side of ['distT', 'distB', 'distL', 'distR']) inline.setAttribute(side, '0');
	const cx = String(Math.max(1, Math.round(image.widthPx * EMU_PER_PIXEL)));
	const cy = String(Math.max(1, Math.round(image.heightPx * EMU_PER_PIXEL)));
	const extent = el(doc, WP_NS, 'wp:extent');
	extent.setAttribute('cx', cx);
	extent.setAttribute('cy', cy);
	const effectExtent = el(doc, WP_NS, 'wp:effectExtent');
	for (const side of ['l', 't', 'r', 'b']) effectExtent.setAttribute(side, '0');
	const docPr = el(doc, WP_NS, 'wp:docPr');
	docPr.setAttribute('id', docPrId);
	docPr.setAttribute('name', `Picture ${docPrId}`);
	if (image.altText) docPr.setAttribute('descr', image.altText);
	if (image.title) docPr.setAttribute('title', image.title);
	const frameLocks = el(doc, WP_NS, 'wp:cNvGraphicFramePr');
	const graphicFrameLocks = el(doc, A_NS, 'a:graphicFrameLocks');
	graphicFrameLocks.setAttribute('noChangeAspect', '1');
	frameLocks.appendChild(graphicFrameLocks);
	const cNvPr = el(doc, PIC_NS, 'pic:cNvPr');
	cNvPr.setAttribute('id', '0');
	cNvPr.setAttribute('name', `Picture ${docPrId}`);
	const nvPicPr = el(doc, PIC_NS, 'pic:nvPicPr');
	appendAll(nvPicPr, cNvPr, el(doc, PIC_NS, 'pic:cNvPicPr'));
	const blip = el(doc, A_NS, 'a:blip');
	blip.setAttributeNS(REL_NS, 'r:embed', relId);
	if (svgRelId) blip.appendChild(svgExtension(doc, svgRelId));
	const stretch = el(doc, A_NS, 'a:stretch');
	stretch.appendChild(el(doc, A_NS, 'a:fillRect'));
	const blipFill = el(doc, PIC_NS, 'pic:blipFill');
	appendAll(blipFill, blip, stretch);
	if (image.pictureFill?.crop) applyPictureCrop(blipFill, image.pictureFill.crop);
	if (image.pictureFill?.effects) applyPictureEffects(blip, image.pictureFill.effects);
	if (image.pictureFill?.layout) applyPictureLayout(blipFill, image.pictureFill.layout);
	const xfrm = el(doc, A_NS, 'a:xfrm');
	const off = el(doc, A_NS, 'a:off');
	off.setAttribute('x', '0');
	off.setAttribute('y', '0');
	const ext = el(doc, A_NS, 'a:ext');
	ext.setAttribute('cx', cx);
	ext.setAttribute('cy', cy);
	appendAll(xfrm, off, ext);
	const prstGeom = el(doc, A_NS, 'a:prstGeom');
	prstGeom.setAttribute('prst', 'rect');
	prstGeom.appendChild(el(doc, A_NS, 'a:avLst'));
	const spPr = el(doc, PIC_NS, 'pic:spPr');
	appendAll(spPr, xfrm, prstGeom);
	const pic = el(doc, PIC_NS, 'pic:pic');
	appendAll(pic, nvPicPr, blipFill, spPr);
	const graphicData = el(doc, A_NS, 'a:graphicData');
	graphicData.setAttribute('uri', PICTURE_GRAPHIC_URI);
	graphicData.appendChild(pic);
	const graphic = el(doc, A_NS, 'a:graphic');
	graphic.appendChild(graphicData);
	appendAll(inline, extent, effectExtent, docPr, frameLocks, graphic);
	drawing.appendChild(inline);
	return drawing;
}

/**
 * Builds or reuses the run for an inline image. Reuses the original drawing XML verbatim
 * (patching only extent/alt text) when the run is recognizably the same picture; synthesizes a
 * minimal, valid `wp:inline` drawing for a newly inserted picture. Refuses to swap one picture's
 * relationship for another in place rather than silently dropping the original XML.
 */
export function createImageRun(
	doc: XmlDocument,
	image: InlineImage,
	base: InlineImage | undefined,
	old: XmlElement | undefined,
	allocator?: RelationshipAllocator,
): XmlElement {
	const node = old ?? makeW(doc, 'r');
	const oldDrawing = findDrawingChild(old);
	const oldContainer = drawingContainer(old);
	if (!base && oldDrawing)
		throw new Error(
			'Cannot safely insert a picture here: this position lines up with a different existing picture. Edit further away from surrounding pictures, then save.',
		);
	if (base && oldDrawing && image.textBoxEditable && base.textBoxEditable) {
		patchTextBox(doc, oldDrawing, image, base);
		clearNonProperties(node, oldContainer);
		return node;
	}
	const diagram = diagramRun(doc, node, image, base, oldDrawing, oldContainer, allocator);
	if (diagram) return diagram;
	const chart = chartRun(doc, node, image, base, oldDrawing, oldContainer, allocator);
	if (chart) return chart;
	// A drawing this model only shows as a placeholder keeps its XML while the rest of the paragraph changes.
	if (base && oldDrawing && image.unsupported && JSON.stringify(image) === JSON.stringify(base)) {
		clearNonProperties(node, oldContainer);
		return node;
	}
	if (base && oldDrawing && image.shape?.editable && base.shape?.editable) {
		const next = rebuildShapeDrawing(doc, image, oldDrawing);
		const choice = oldDrawing.parentNode as XmlElement | null;
		if (oldContainer?.localName === 'AlternateContent' && choice) {
			// Keep the wrapper and its fallback: only the modern drawing is rebuilt.
			choice.replaceChild(next, oldDrawing);
			patchFallbackSize(oldContainer, image, base);
			clearNonProperties(node, oldContainer);
			return node;
		}
		clearNonProperties(node);
		node.appendChild(next);
		return node;
	}
	if (!base && image.shape?.editable && !image.partName) {
		if (!allocator) throw new Error('Cannot insert a shape without an id allocator.');
		const id = allocator.docPrIds.next();
		clearNonProperties(node);
		node.appendChild(
			buildShapeDrawing(
				doc,
				image,
				{ id, ...(image.altText ? { descr: image.altText } : {}) },
				251658240 + (Number(id) % 4096) * 1024,
			),
		);
		return node;
	}
	if (image.watermark) {
		clearNonProperties(node);
		node.appendChild(
			buildWatermarkRun(doc, image.watermark, allocator?.docPrIds.next() ?? '1').lastChild!,
		);
		return node;
	}
	if (!base && image.textBoxEditable && !image.partName) {
		if (!allocator) throw new Error('Cannot insert a text box without an id allocator.');
		clearNonProperties(node);
		node.appendChild(buildTextBoxDrawing(doc, image, allocator.docPrIds.next()));
		return node;
	}
	if (base && oldDrawing && sameUnderlyingImage(base, image)) {
		if (
			image.pictureFill &&
			JSON.stringify(image.pictureFill) !== JSON.stringify(base.pictureFill)
		) {
			const fill = oldDrawing.getElementsByTagNameNS(PIC_NS, 'blipFill')[0];
			const blip = fill?.getElementsByTagNameNS(A_NS, 'blip')[0];
			if (fill && image.pictureFill.crop !== undefined)
				applyPictureCrop(fill, image.pictureFill.crop);
			if (blip && image.pictureFill.effects !== undefined)
				applyPictureEffects(blip, image.pictureFill.effects);
			if (fill && image.pictureFill.layout !== undefined)
				applyPictureLayout(fill, image.pictureFill.layout);
		}
		if (image.widthPx !== base.widthPx || image.heightPx !== base.heightPx)
			updateExtent(oldDrawing, image.widthPx, image.heightPx);
		if (image.altText !== base.altText || image.title !== base.title)
			updateDocPr(oldDrawing, image.altText, image.title);
		if (
			image.anchored &&
			image.placement &&
			JSON.stringify(image.placement) !== JSON.stringify(base.placement)
		)
			updatePlacement(oldDrawing, doc, image.placement);
		clearNonProperties(node, oldContainer);
		return node;
	}
	if (base)
		throw new Error(
			'Cannot replace an existing picture with a different one in place; remove it and insert a new picture instead.',
		);
	if (!allocator) throw new Error('Cannot insert a new picture without a relationship allocator.');
	const relId = allocator.addImage(image.partName);
	const svgRelId = image.svgPartName ? allocator.addImage(image.svgPartName) : undefined;
	const drawing = buildInlineDrawing(doc, image, relId, allocator.docPrIds.next(), svgRelId);
	clearNonProperties(node);
	node.appendChild(drawing);
	return node;
}
