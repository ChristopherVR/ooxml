import type { PptxElement, PptxSlide, ZoomPptxElement } from '../../types';

type Preview = Pick<
	ZoomPptxElement,
	'targetSlideIndex' | 'zoomImageType' | 'zoomShowBackground' | 'imagePath' | 'imageData'
>;

function* zoomPreviews(elements: readonly PptxElement[]): Iterable<Preview> {
	for (const element of elements) {
		if (element.type === 'group') yield* zoomPreviews(element.children);
		if (element.type === 'zoom') {
			if (element.zoomType === 'summary') yield* element.summaryTargets ?? [];
			else yield element;
		}
	}
}

/** Refresh edited slide previews before the writer fingerprints their owning slides. */
export async function refreshZoomPreviews(
	slides: readonly PptxSlide[],
	changed: (slide: PptxSlide) => boolean,
	render: (slide: PptxSlide, transparent: boolean) => Promise<string | undefined>,
): Promise<void> {
	const dirty = new Map<PptxSlide, boolean>();
	const images = new Map<string, string | undefined>();
	const replacements: { owner: PptxSlide; preview: Preview; png: string }[] = [];
	const isChanged = (slide: PptxSlide) => {
		const previous = dirty.get(slide);
		if (previous !== undefined) return previous;
		const result = changed(slide);
		dirty.set(slide, result);
		return result;
	};
	for (const owner of slides) {
		for (const preview of zoomPreviews(owner.elements)) {
			if (preview.zoomImageType === 'cover') continue;
			const target = slides[preview.targetSlideIndex];
			if (!target || (preview.imageData && !isChanged(target) && !isChanged(owner))) continue;
			const transparent = preview.zoomShowBackground === false;
			const key = `${preview.targetSlideIndex}:${transparent}`;
			if (!images.has(key)) {
				images.set(key, await render(target, transparent).catch(() => undefined));
			}
			const png = images.get(key);
			if (png?.startsWith('data:image/png;base64,') && png !== preview.imageData) {
				replacements.push({ owner, preview, png });
			}
		}
	}
	for (const { owner, preview, png } of replacements) {
		preview.imageData = png;
		preview.imagePath = undefined;
		owner.isDirty = true;
	}
}
