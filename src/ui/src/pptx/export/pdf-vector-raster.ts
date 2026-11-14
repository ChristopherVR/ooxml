/**
 * Draw one element of a mounted slide as a picture, on its own: everything else on the stage
 * is hidden (`visibility`) for the capture, so the picture holds only that element (or only its
 * box, or only its text shadow) over a transparent background and can sit between vector items.
 * The capture uses the same raster pipeline as the picture export (`foreignObject` first, then
 * the binding's html2canvas driver); the live stage is put back exactly as it was.
 */
import { drawableToPdfImage } from '../../export/pdf-image';
import type { PdfImageResource } from '../../export/pdf-vector';
import type {
	FontStyleDocumentLike,
	LinkStyleDocumentLike,
} from '../render/foreign-object-font-embed';
import { buildForeignObjectSvgBody } from '../render/foreign-object-svg-document';
import type { RasterSourceRect, RasterStrategyOptions } from './rasterize-element-strategy';
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
	const saved: Saved[] = [];
	try {
		if (target !== stage) setStyle(saved, stage, 'visibility', 'hidden');
		setStyle(saved, target, 'visibility', 'visible');
		if (mode !== 'subtree')
			for (const child of Array.from(target.children))
				setStyle(saved, child, 'visibility', 'hidden');
		if (mode === 'shadow') {
			setStyle(saved, target, 'color', 'transparent');
			setStyle(saved, target, '-webkit-text-fill-color', 'transparent');
			setStyle(saved, target, 'background', 'none');
			setStyle(saved, target, 'border-color', 'transparent');
			setStyle(saved, target, 'box-shadow', 'none');
			setStyle(saved, target, 'text-decoration-color', 'transparent');
		}
		const body = await buildForeignObjectSvgBody(
			stage,
			doc as Document & FontStyleDocumentLike & LinkStyleDocumentLike,
			{ width: naturalWidth, height: naturalHeight },
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
	} catch {
		return undefined;
	} finally {
		restore(saved);
	}
}
