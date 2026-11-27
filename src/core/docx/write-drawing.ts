import { setSvgPictureRelationship } from '../drawingml/svg-picture';
import { applyPictureCrop } from '../drawingml/picture-crop';
import { applyPictureEffects } from '../drawingml/picture-effects';
import { applyPictureLayout } from '../drawingml/picture-layout';
import { MODEL3D_REL } from '../drawingml/model3d';
import { RELATIONSHIP_TYPES } from '../opc/index';
// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Writes inline picture runs: reuses original wp:inline/pic:pic XML verbatim (only patching
// extent/alt text) for unchanged pictures, and synthesizes a minimal valid drawing for insertions.
import type { InlineImage } from './model';
import { EMU_PER_PIXEL } from './drawing';
import { buildWatermarkRun } from './watermark';
import { buildPictureWatermarkRun } from './picture-watermark';
import { buildSignatureLineRun } from './signature-line';
import { applyWebVideo } from './web-video';
import { buildTextBoxDrawing, patchTextBox } from './write-text-box';
import { buildShapeDrawing, rebuildShapeDrawing } from './write-shape';
import { chartRun } from './write-chart-drawing';
import { buildInkDrawing, patchInkXfrm } from './write-ink-drawing';
import {
	clearNonProperties,
	updateDocPr,
	updateExtent,
	updatePlacement,
	updateDrawingTransform,
	updateTextBoxDirection,
} from './write-drawing-patch';
import type { RelationshipAllocator } from './relationship-allocator';
import { patchFallbackSize } from './write-fallback';
import {
	patchDrawingOrder,
	patchDrawingWrap,
	placementPositionChanged,
} from './write-drawing-wrap';
import { diagramRun } from './write-diagram-drawing';
import { isOleDrawing } from './drawing-payload';
import {
	model3dAlternateContent,
	officePackageDrawing,
	updateOleIconSize,
} from './write-drawing-payload';
import { isElement, makeNS, makeW, named, REL_NS, type XmlDocument, type XmlElement } from './xml';
import { patchVmlFloat } from './vml-placement';

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
		if (
			isElement(child) &&
			(named(child, 'drawing') || named(child, 'pict') || isOleDrawing(child))
		)
			return child;
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
		if (named(child, 'drawing') || named(child, 'pict') || isOleDrawing(child)) return child;
		if (child.localName === 'AlternateContent') {
			const inner = alternateContentDrawing(child);
			if (inner) return inner;
		}
	}
	return undefined;
}

function sameUnderlyingImage(a: InlineImage, b: InlineImage): boolean {
	return (
		Boolean(a.partName) &&
		a.relId === b.relId &&
		a.partName === b.partName &&
		a.payload?.kind === b.payload?.kind &&
		a.payload?.partName === b.payload?.partName &&
		(a.payload?.kind !== 'officePackage' ||
			(b.payload?.kind === 'officePackage' && a.payload.progId === b.payload.progId))
	);
}

function el(doc: XmlDocument, ns: string, qualified: string): XmlElement {
	return makeNS(doc, ns, qualified);
}
/** @xmldom/xmldom elements only implement `appendChild`, not the newer `append(...)` convenience. */
function appendAll(parent: XmlElement, ...items: XmlElement[]): void {
	for (const item of items) parent.appendChild(item);
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
	docPr.setAttribute('name', image.drawingName ?? `Picture ${docPrId}`);
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
	if (svgRelId) setSvgPictureRelationship(blip, svgRelId);
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
function patchArrangement(
	doc: XmlDocument,
	drawing: XmlElement,
	image: InlineImage,
	base?: InlineImage,
): void {
	if (
		image.altText !== base?.altText ||
		image.title !== base?.title ||
		image.drawingName !== base?.drawingName
	)
		updateDocPr(drawing, image.altText, image.title, image.drawingName);
	patchDrawingWrap(drawing, doc, image, base);
	if (
		image.anchored &&
		image.placement &&
		base?.anchored &&
		placementPositionChanged(image.placement, base.placement)
	)
		updatePlacement(drawing, doc, image.placement, base.placement);
	patchDrawingOrder(drawing, image.placement, base?.placement);
	const transform = image.shape ?? image.transform;
	const previous = base?.shape ?? base?.transform;
	if (
		transform?.rotation !== previous?.rotation ||
		transform?.flipH !== previous?.flipH ||
		transform?.flipV !== previous?.flipV
	)
		updateDrawingTransform(drawing, transform ?? {});
	if (image.textBoxContent?.vertical !== base?.textBoxContent?.vertical)
		updateTextBoxDirection(drawing, image.textBoxContent?.vertical);
}

function shapeArtwork(image: InlineImage): string {
	if (!image.shape) return '';
	const { rotation: _rot, flipH: _h, flipV: _v, ...artwork } = image.shape;
	return JSON.stringify(artwork);
}

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
		patchArrangement(doc, oldDrawing, image, base);
		clearNonProperties(node, oldContainer);
		return node;
	}
	const diagram = diagramRun(doc, node, image, base, oldDrawing, oldContainer, allocator);
	if (diagram) return diagram;
	const chart = chartRun(doc, node, image, base, oldDrawing, oldContainer, allocator);
	if (chart) return chart;
	// Kept ink keeps its drawing (its part is saved by ink-save.ts); new ink gets a part on save.
	if (base?.ink && oldDrawing && image.ink?.relId && image.ink.relId === base.ink.relId) {
		if (image.widthPx !== base.widthPx || image.heightPx !== base.heightPx) {
			updateExtent(oldDrawing, image.widthPx, image.heightPx);
			patchInkXfrm(oldDrawing, image.widthPx, image.heightPx);
		}
		patchArrangement(doc, oldDrawing, image, base);
		clearNonProperties(node, oldContainer);
		return node;
	}
	// Ink with no drawing of its own here (new, pasted, moved to another paragraph or copied)
	// gets a drawing and a part of its own, so no two drawings share one part.
	if (!base && image.ink) {
		if (!allocator) throw new Error('Cannot insert ink without a relationship allocator.');
		const relId = allocator.addInk(image.ink);
		clearNonProperties(node);
		node.appendChild(buildInkDrawing(doc, image, relId, allocator.docPrIds.next(image.drawingId)));
		return node;
	}
	if (image.payload && !base) {
		if (!allocator)
			throw new Error('Cannot insert a native object without a relationship allocator.');
		const previewId = allocator.addImage(image.partName);
		const id = allocator.docPrIds.next();
		const payload = image.payload;
		const relType = payload.kind === 'model3d' ? MODEL3D_REL : RELATIONSHIP_TYPES.package;
		const payloadId = allocator.addBinaryPart(relType, payload.partName);
		clearNonProperties(node);
		node.appendChild(
			payload.kind === 'model3d'
				? model3dAlternateContent(
						doc,
						buildInlineDrawing(doc, image, previewId, id),
						payloadId,
						previewId,
						image,
					)
				: officePackageDrawing(doc, image, payload.progId, payloadId, previewId, id),
		);
		return node;
	}
	// A drawing this model only shows as a placeholder keeps its XML while the rest of the paragraph changes.
	if (base && oldDrawing && image.unsupported && JSON.stringify(image) === JSON.stringify(base)) {
		clearNonProperties(node, oldContainer);
		return node;
	}
	if (
		base?.shape &&
		oldDrawing &&
		image.shape &&
		Boolean(image.shape.groupChildren) !== Boolean(base.shape.groupChildren)
	) {
		if (!allocator) throw new Error('Cannot replace a drawing without an id allocator.');
		clearNonProperties(node);
		node.appendChild(
			buildShapeDrawing(doc, image, { id: allocator.docPrIds.next(image.drawingId) }),
		);
		return node;
	}
	if (
		base &&
		oldDrawing &&
		image.shape &&
		base.shape &&
		shapeArtwork(image) === shapeArtwork(base) &&
		image.widthPx === base.widthPx &&
		image.heightPx === base.heightPx
	) {
		patchArrangement(doc, oldDrawing, image, base);
		clearNonProperties(node, oldContainer);
		return node;
	}
	if (base && oldDrawing && image.shape?.editable && base.shape?.editable) {
		const next = rebuildShapeDrawing(doc, image, oldDrawing);
		if (
			image.altText !== base.altText ||
			image.title !== base.title ||
			image.drawingName !== base.drawingName
		)
			updateDocPr(next, image.altText, image.title, image.drawingName);
		if (image.textBoxContent?.vertical !== base.textBoxContent?.vertical)
			updateTextBoxDirection(next, image.textBoxContent?.vertical);
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
	if (!base && (image.shape?.editable || image.shape?.groupChildren) && !image.partName) {
		if (!allocator) throw new Error('Cannot insert a shape without an id allocator.');
		const id = allocator.docPrIds.next(image.drawingId);
		clearNonProperties(node);
		node.appendChild(
			buildShapeDrawing(
				doc,
				image,
				{
					id,
					...(image.altText ? { descr: image.altText } : {}),
					...(image.title ? { title: image.title } : {}),
				},
				251658240 + (Number(id) % 4096) * 1024,
			),
		);
		return node;
	}
	if (image.pictureWatermark) {
		if (!allocator)
			throw new Error('Cannot insert a picture watermark without a relationship allocator.');
		const spec = image.pictureWatermark;
		clearNonProperties(node);
		node.appendChild(
			buildPictureWatermarkRun(
				doc,
				spec,
				allocator.addImage(spec.partName),
				allocator.docPrIds.next(),
			).lastChild!,
		);
		return node;
	}
	if (!base && image.signatureLine) {
		if (!allocator)
			throw new Error('Cannot insert a signature line without a relationship allocator.');
		clearNonProperties(node);
		node.appendChild(
			buildSignatureLineRun(
				doc,
				image.signatureLine,
				allocator.addImage(image.partName),
				allocator.docPrIds.next(),
				{
					widthPt: Math.round(image.widthPx * 75) / 100,
					heightPt: Math.round(image.heightPx * 75) / 100,
				},
			).lastChild!,
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
		const drawing = buildTextBoxDrawing(doc, image, allocator.docPrIds.next());
		patchArrangement(doc, drawing, image);
		node.appendChild(drawing);
		return node;
	}
	if (base && oldDrawing && sameUnderlyingImage(base, image)) {
		if (named(oldDrawing, 'pict')) {
			if (base.vmlFloat) patchVmlFloat(oldDrawing, image, base);
			clearNonProperties(node, oldContainer);
			return node;
		}
		if (
			image.pictureFill &&
			JSON.stringify(image.pictureFill) !== JSON.stringify(base.pictureFill)
		) {
			const fill = oldDrawing.getElementsByTagNameNS(PIC_NS, 'blipFill')[0];
			const blip = fill?.getElementsByTagNameNS(A_NS, 'blip')[0];
			if (fill && (image.pictureFill.crop !== undefined || base.pictureFill?.crop))
				applyPictureCrop(fill, image.pictureFill.crop);
			if (blip && image.pictureFill.effects !== undefined)
				applyPictureEffects(blip, image.pictureFill.effects);
			if (fill && image.pictureFill.layout !== undefined)
				applyPictureLayout(fill, image.pictureFill.layout);
		}
		if (image.widthPx !== base.widthPx || image.heightPx !== base.heightPx) {
			if (isOleDrawing(oldDrawing)) updateOleIconSize(oldDrawing, image.widthPx, image.heightPx);
			else
				updateExtent(
					image.payload && oldContainer ? oldContainer : oldDrawing,
					image.widthPx,
					image.heightPx,
				);
		}
		patchArrangement(doc, oldDrawing, image, base);
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
	if (image.webVideo) {
		const docPr = drawing.getElementsByTagNameNS(WP_NS, 'docPr')[0];
		if (docPr) {
			docPr.setAttribute('name', `Online Media ${docPr.getAttribute('id')}`);
			applyWebVideo(doc, docPr, image.webVideo);
		}
	}
	patchArrangement(doc, drawing, image);
	clearNonProperties(node);
	node.appendChild(drawing);
	return node;
}
