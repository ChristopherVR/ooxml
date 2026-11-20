import {
	applyImagePathPatches,
	collectImagePaths,
	resolveTableCellImageUrls,
	resolveTableStyleImageUrls,
	resolveTextFillBlipUrls,
} from 'ooxml-ui/pptx';
import type { ParsedTableStyleMap, PptxHandler, PptxSlide } from 'pptx-viewer-core';

/**
 * Pure async helpers for the load pipeline, ported from the Vue binding's
 * `useLoadContent` body. Media/image *collection* comes from
 * `pptx-viewer-shared`; these functions add the handler-driven URL
 * resolution and immutable slide patching.
 */

/** Revoke any `blob:` URLs in the list (no-op for data URLs). */
export function revokeBlobUrls(urls: Iterable<string>): void {
	for (const url of urls) {
		if (url.startsWith('blob:')) {
			URL.revokeObjectURL(url);
		}
	}
}

/**
 * Resolve lazily-loaded picture URLs and patch them into the slide tree
 * immutably (groups recursed). Returns the input array when nothing changed.
 */
export async function resolveLazyImages(
	handler: PptxHandler,
	slides: PptxSlide[],
): Promise<PptxSlide[]> {
	const { paths, refs } = collectImagePaths(slides);
	if (paths.size === 0) {
		return slides;
	}

	const resolvedMap = new Map<string, string>();
	await Promise.all(
		Array.from(paths).map(async (path) => {
			try {
				const url = await handler.getImageData(path);
				if (url) {
					resolvedMap.set(path, url);
				}
			} catch {
				// Non-critical: image will show as broken.
			}
		}),
	);

	return slides.map((slide) => {
		const newElements = applyImagePathPatches(slide.elements, resolvedMap, refs);
		return newElements === slide.elements ? slide : { ...slide, elements: newElements };
	});
}

/**
 * Resolve lazily-loaded table cell image-fill URLs (`a:tcPr/a:blipFill`) and
 * patch them into the slide tree immutably. Same lazy-load story as
 * {@link resolveLazyImages}, but for a cell's `backgroundImageFillPath`
 * rather than a top-level element field.
 */
export function resolveLazyTableCellImages(
	handler: PptxHandler,
	slides: PptxSlide[],
): Promise<PptxSlide[]> {
	return resolveTableCellImageUrls(slides, (path) => handler.getImageData(path));
}

/**
 * Resolve lazily-loaded whole-table-STYLE image-fill URLs
 * (`a:tcStyle/a:fill/a:blipFill` on `ppt/tableStyles.xml`) and patch them
 * into the table style map immutably. Same lazy-load story as
 * {@link resolveLazyTableCellImages}, but for a presentation-level style
 * section fill rather than a per-cell one.
 */
export function resolveLazyTableStyleImages(
	handler: PptxHandler,
	tableStyleMap: ParsedTableStyleMap | undefined,
): Promise<ParsedTableStyleMap | undefined> {
	return resolveTableStyleImageUrls(tableStyleMap, (path) => handler.getImageData(path));
}

/**
 * Resolve lazily-loaded text-run picture-fill URLs (`a:rPr > a:blipFill`,
 * COM-verified regression: `audit-text` slide 13's "PICTURE FILL" run) and
 * patch them into the slide tree immutably. Same lazy-load story as
 * {@link resolveLazyTableCellImages}, but for a run's `textFillBlipUrl`.
 */
export function resolveLazyTextFillBlips(
	handler: PptxHandler,
	slides: PptxSlide[],
): Promise<PptxSlide[]> {
	return resolveTextFillBlipUrls(slides, (path) => handler.getImageData(path));
}
