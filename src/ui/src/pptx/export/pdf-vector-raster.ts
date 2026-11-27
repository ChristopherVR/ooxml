/**
 * Draw one element of a mounted slide as a picture, on its own: everything else on the stage
 * is hidden (`visibility`) for the capture, so the picture holds only that element (or only its
 * box, or only its text shadow) over a transparent background and can sit between vector items.
 * The capture uses the same raster pipeline as the picture export (`foreignObject` first, then
 * the binding's html2canvas driver); the live stage is put back exactly as it was.
 */
import { drawableToPdfImage } from '../../export/pdf-image';
import type { PdfImageResource } from 'ooxml-core/pdf/write';
import type {
	ExternalFontFaceCssResult,
	FontStyleDocumentLike,
	LinkStyleDocumentLike,
} from '../render/foreign-object-font-embed';
import { buildForeignObjectSvgBody } from '../render/foreign-object-svg-document';
import type { RasterSourceRect, RasterStrategyOptions } from './rasterize-element-strategy';
import { RASTER_TIMEOUT_MS, settleWithin } from './pdf-vector-timeout';
import { rasterizeWindow } from './rasterize-element-strategy';

/**
 * `subtree` draws the element and its descendants; `self` its own box without its children;
 * `shadow` only the shadow of its own text (glyphs and box made transparent).
 */
export type IsolatedRasterMode = 'subtree' | 'self' | 'shadow';

export interface IsolatedRasterOptions {
	/** Device pixels per CSS px. */
	scale: number;
	html2canvasFallback?: RasterStrategyOptions['html2canvasFallback'];
	/** The document's external font CSS, when collected once for a whole export. */
	externalFonts?: () => Promise<ExternalFontFaceCssResult>;
}

interface Saved {
	el: HTMLElement | SVGElement;
	property: string;
	value: string;
	priority: string;
}

function setStyle(saved: Saved[], el: Element, property: string, value: string): void {
	const styled = el as HTMLElement;
	if (!styled.style) return;
	saved.push({
		el: styled,
		property,
		value: styled.style.getPropertyValue(property),
		priority: styled.style.getPropertyPriority(property),
	});
	styled.style.setProperty(property, value, 'important');
}

function restore(saved: Saved[]): void {
	for (const { el, property, value, priority } of saved.reverse()) {
		if (value) el.style.setProperty(property, value, priority);
		else el.style.removeProperty(property);
	}
}

/**
 * Isolate `target` on the live stage the way a capture of it in `mode` sees it (everything else
 * hidden; for `self` and `shadow` its children too; for `shadow` its glyphs and box transparent).
 * Returns the function that puts the stage back exactly as it was.
 */
export function isolateLive(
	stage: HTMLElement,
	target: Element,
	mode: IsolatedRasterMode,
): () => void {
	const saved: Saved[] = [];
	try {
		hideAllBut(saved, stage, target, mode);
	} catch (error) {
		restore(saved);
		throw error;
	}
	return () => restore(saved);
}

function hideAllBut(
	saved: Saved[],
	stage: HTMLElement,
	target: Element,
	mode: IsolatedRasterMode,
): void {
	if (target !== stage) setStyle(saved, stage, 'visibility', 'hidden');
	setStyle(saved, target, 'visibility', 'visible');
	if (mode !== 'subtree')
		for (const child of Array.from(target.children)) setStyle(saved, child, 'visibility', 'hidden');
	// A `visibility` an element declares in its own style beats the hidden stage it inherits
	// from (renderers declare `visible` on every shape), so each one outside what the picture
	// draws is hidden too.
	for (const el of Array.from(stage.querySelectorAll('*'))) {
		if (el === target || el.parentElement === target) continue;
		if (mode === 'subtree' && target.contains(el)) continue;
		const declared = (el as HTMLElement).style?.getPropertyValue('visibility');
		if (declared && declared !== 'hidden') setStyle(saved, el, 'visibility', 'hidden');
	}
	if (mode === 'shadow') {
		setStyle(saved, target, 'color', 'transparent');
		setStyle(saved, target, '-webkit-text-fill-color', 'transparent');
		// Longhands only: a shorthand reads back empty when only a longhand was set inline, so
		// restoring it would drop the live element's own value.
		setStyle(saved, target, 'background-color', 'transparent');
		setStyle(saved, target, 'background-image', 'none');
		for (const side of ['top', 'right', 'bottom', 'left'])
			setStyle(saved, target, `border-${side}-color`, 'transparent');
		setStyle(saved, target, 'box-shadow', 'none');
		setStyle(saved, target, 'text-decoration-color', 'transparent');
	}
}

const noFallback: RasterStrategyOptions['html2canvasFallback'] = () =>
	Promise.reject(new Error('No picture fallback for this capture.'));

/**
 * Capture `target` within `rect` (stage px, as painted). `undefined` when no strategy could
 * produce readable pixels.
 */
export async function rasterizeIsolated(
	stage: HTMLElement,
	target: Element,
	mode: IsolatedRasterMode,
	rect: RasterSourceRect,
	naturalWidth: number,
	naturalHeight: number,
	options: IsolatedRasterOptions,
): Promise<PdfImageResource | undefined> {
	const doc = stage.ownerDocument;
	let restoreLive = () => {};
	try {
		restoreLive = isolateLive(stage, target, mode);
		// Bounded, so a stalled font or image fetch cannot hang the export; the finally below then
		// restores the live stage either way.
		const work = (async () => {
			const body = await buildForeignObjectSvgBody(
				stage,
				doc as Document & FontStyleDocumentLike & LinkStyleDocumentLike,
				{
					width: naturalWidth,
					untransformedRoot: true,
					height: naturalHeight,
					...(options.externalFonts ? { externalFonts: options.externalFonts() } : {}),
				},
			);
			const output = {
				width: Math.max(1, Math.round(rect.width * options.scale)),
				height: Math.max(1, Math.round(rect.height * options.scale)),
			};
			const { canvas } = await rasterizeWindow(body, rect, output, undefined, {
				html2canvasFallback: options.html2canvasFallback ?? noFallback,
				onStrategyFailed: () => undefined,
			});
			return await drawableToPdfImage(doc, canvas, canvas.width, canvas.height, 8192);
		})();
		return await settleWithin(work, RASTER_TIMEOUT_MS, undefined);
	} catch {
		return undefined;
	} finally {
		restoreLive();
	}
}
