/**
 * The pieces of the vector PDF collector (`pdf-vector-collect.ts`) that read one element:
 * why it needs a picture, its content box, the bounds a picture of it covers, and its inline
 * SVG as a display item. The diagnostic codes of each reason live here too.
 */
import { collectSvgImages, drawSvgSubtree } from 'ooxml-core/pdf/write';
import { PdfFallback } from 'ooxml-core/pdf/write';
import { PdfVectorPage, type PdfImageResource } from 'ooxml-core/pdf/write';
import type { PdfExportDiagnosticCode } from './pdf-export-diagnostics';
import { SHADOW_REASON } from './pdf-vector-box';
import { roundedRectPath } from './pdf-vector-css';
import type { ImageItem, Matrix, PdfClip, SvgItem } from './pdf-vector-display-list';
import { multiply } from './pdf-vector-display-list';
import { normalizeInlineSvg, viewBoxMatrix } from './pdf-vector-svg';

export const FILTER_REASON = 'filterAsPicture' satisfies PdfExportDiagnosticCode;
export const BLEND_REASON = 'blendAsPicture' satisfies PdfExportDiagnosticCode;
export const THREE_D_REASON = 'threeDAsPicture' satisfies PdfExportDiagnosticCode;
export const MEDIA_REASON = 'mediaAsPicture' satisfies PdfExportDiagnosticCode;
export const TEXT_SHADOW_REASON = 'textShadowAsPicture' satisfies PdfExportDiagnosticCode;
export const SEARCHABLE_NOTE = 'searchableText' satisfies PdfExportDiagnosticCode;
export const OPACITY_NOTE = 'groupOpacity' satisfies PdfExportDiagnosticCode;
export const SVG_REASON = 'svgAsPicture' satisfies PdfExportDiagnosticCode;
const FILL_TEXT_REASON = 'fillTextAsPicture' satisfies PdfExportDiagnosticCode;

export const MEDIA_TAGS = new Set(['canvas', 'video']);
const OPAQUE_TAGS = new Set(['iframe', 'object', 'embed', 'math']);

export interface Box {
	x: number;
	y: number;
	w: number;
	h: number;
}

export const px = (value: string) => Number.parseFloat(value) || 0;

/** An image item whose picture is decoded or captured later. */
export function placeholderImage(
	matrix: Matrix,
	clips: readonly PdfClip[],
	alpha: number,
	box: Box,
): ImageItem {
	return {
		kind: 'image',
		matrix,
		clips,
		alpha,
		x: box.x,
		y: box.y,
		width: box.w,
		height: box.h,
		image: { width: 0, height: 0, data: new Uint8Array() },
	};
}

/** The stage-px rectangle `box` covers once painted with `matrix`, padded and kept on the slide. */
export function paintedBounds(
	matrix: Matrix,
	box: Box,
	pad: number,
	width: number,
	height: number,
): Box | undefined {
	const corners = [
		[box.x, box.y],
		[box.x + box.w, box.y],
		[box.x, box.y + box.h],
		[box.x + box.w, box.y + box.h],
	].map(([cx, cy]) => [
		matrix[0] * cx! + matrix[2] * cy! + matrix[4],
		matrix[1] * cx! + matrix[3] * cy! + matrix[5],
	]);
	const x0 = Math.max(0, Math.floor(Math.min(...corners.map((c) => c[0]!)) - pad));
	const y0 = Math.max(0, Math.floor(Math.min(...corners.map((c) => c[1]!)) - pad));
	const x1 = Math.min(width, Math.ceil(Math.max(...corners.map((c) => c[0]!)) + pad));
	const y1 = Math.min(height, Math.ceil(Math.max(...corners.map((c) => c[1]!)) + pad));
	return x1 > x0 && y1 > y0 ? { x: x0, y: y0, w: x1 - x0, h: y1 - y0 } : undefined;
}

/** How far past its box an effect's picture reaches (blur and shadow spread). */
export const effectPad = (reason: PdfExportDiagnosticCode, box: Box): number =>
	reason === FILTER_REASON || reason === THREE_D_REASON || reason === SHADOW_REASON
		? Math.min(200, 24 + 0.25 * Math.max(box.w, box.h))
		: 2;

/** Why the element and its subtree must be a picture (`undefined` when they can be vector). */
export function subtreeRasterReason(
	el: Element,
	style: CSSStyleDeclaration,
	clipReadable: boolean,
): PdfExportDiagnosticCode | undefined {
	if (style.filter && style.filter !== 'none') return FILTER_REASON;
	if ((style.mixBlendMode && style.mixBlendMode !== 'normal') || hasMask(style) || !clipReadable)
		return BLEND_REASON;
	if (OPAQUE_TAGS.has(el.localName) || (el as HTMLElement).shadowRoot) return MEDIA_REASON;
	if (
		style.backgroundClip === 'text' ||
		style.getPropertyValue('-webkit-background-clip') === 'text'
	)
		return FILL_TEXT_REASON;
	return undefined;
}

function hasMask(style: CSSStyleDeclaration): boolean {
	const mask = style.getPropertyValue('mask-image') || style.getPropertyValue('-webkit-mask-image');
	const backdrop =
		style.getPropertyValue('backdrop-filter') || style.getPropertyValue('-webkit-backdrop-filter');
	return (!!mask && mask !== 'none') || (!!backdrop && backdrop !== 'none');
}

export const hasText = (el: Element) => !!el.textContent?.trim();
export const hasOwnText = (el: Element) =>
	Array.from(el.childNodes).some((node) => node.nodeType === 3 && !!(node as Text).data.trim());

/** The padding box (inside the borders), where `overflow` clips children. */
export function paddingBox(style: CSSStyleDeclaration, box: Box): Box {
	return {
		x: box.x + px(style.borderLeftWidth),
		y: box.y + px(style.borderTopWidth),
		w: box.w - px(style.borderLeftWidth) - px(style.borderRightWidth),
		h: box.h - px(style.borderTopWidth) - px(style.borderBottomWidth),
	};
}

/** The content box (inside borders and padding). */
export function contentBox(style: CSSStyleDeclaration, box: Box): Box {
	const inner = paddingBox(style, box);
	return {
		x: inner.x + px(style.paddingLeft),
		y: inner.y + px(style.paddingTop),
		w: inner.w - px(style.paddingLeft) - px(style.paddingRight),
		h: inner.h - px(style.paddingTop) - px(style.paddingBottom),
	};
}

/** A clip to `box` in the space `matrix` maps to the stage. */
export const boxClip = (matrix: Matrix, box: Box): PdfClip => ({
	matrix,
	d: roundedRectPath(box.x, box.y, box.w, box.h),
});

/** Whether `el` is a cell of a `border-collapse: collapse` table. */
export function isCollapsedCell(el: Element, style: CSSStyleDeclaration, view: Window): boolean {
	const table = style.display === 'table-cell' ? el.closest('table') : null;
	return !!table && view.getComputedStyle(table).borderCollapse === 'collapse';
}

/**
 * An inline `<svg>` as a display item, or `undefined` when the shared converter cannot write
 * it (filters, masks, patterns, foreign objects, or anything it reports as a fallback).
 */
export function inlineSvgItem(
	el: SVGSVGElement,
	style: CSSStyleDeclaration,
	matrix: Matrix,
	clips: readonly PdfClip[],
	alpha: number,
	content: Box,
): { item: SvgItem; hrefs: string[]; images: Map<string, PdfImageResource> } | undefined {
	const copy = normalizeInlineSvg(el);
	if (copy.querySelector('[filter], [mask], foreignObject, pattern')) return undefined;
	const hrefs = collectSvgImages(copy);
	try {
		const stand = { width: 1, height: 1, data: new Uint8Array(3) };
		drawSvgSubtree(new PdfVectorPage(10, 10), copy, new Map(hrefs.map((href) => [href, stand])));
	} catch (error) {
		if (error instanceof PdfFallback) return undefined;
		throw error;
	}
	const images = new Map<string, PdfImageResource>();
	return {
		item: {
			kind: 'svg',
			matrix: multiply(
				matrix,
				multiply([1, 0, 0, 1, content.x, content.y], viewBoxMatrix(el, content.w, content.h)),
			),
			clips: style.overflow === 'visible' ? clips : [...clips, boxClip(matrix, content)],
			alpha,
			svg: copy,
			images,
		},
		hrefs,
		images,
	};
}

/** Fit a decoded picture into its box like `object-fit` / `background-size`. */
export function fitImage(item: ImageItem, image: PdfImageResource, fit: 'cover' | 'contain'): void {
	const ratio = image.width / image.height;
	if (!(ratio > 0)) return;
	const boxRatio = item.width / item.height;
	const wide = fit === 'cover' ? ratio > boxRatio : ratio < boxRatio;
	const w = wide ? item.height * ratio : item.width;
	const h = wide ? item.height : item.width / ratio;
	item.x += (item.width - w) / 2;
	item.y += (item.height - h) / 2;
	item.width = w;
	item.height = h;
}
