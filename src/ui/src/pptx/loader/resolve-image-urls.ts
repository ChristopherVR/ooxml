import type { PptxHandler, PptxSlide } from 'ooxml-core/pptx';

import { applyImagePathPatches } from './element-patch-walker';
import { collectImagePaths } from './load-content-helpers';

/** Resolve lazily-loaded picture URLs and patch them into the slide tree. */
export async function resolveImageUrls(
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
