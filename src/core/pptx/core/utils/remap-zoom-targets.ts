import type { PptxElement, PptxSlide } from '../types';

/** Rebase positional Zoom targets onto a new order while keeping their slide or section identity. */
export function remapZoomTargets(before: readonly PptxSlide[], after: PptxSlide[]): PptxSlide[] {
	const bySlide = new Map(after.map(({ id }, index) => [id, index]));
	const bySection = new Map<string, number>();
	for (const [index, slide] of after.entries()) {
		if (slide.sectionId && !bySection.has(slide.sectionId)) bySection.set(slide.sectionId, index);
	}
	const resolve = (index: number, sectionId?: string): number => {
		const section = sectionId ? bySection.get(sectionId) : undefined;
		const slide = before[index];
		return section ?? (slide ? bySlide.get(slide.id) : undefined) ?? index;
	};
	const remap = (element: PptxElement): PptxElement => {
		if (element.type === 'group') {
			const children = element.children.map(remap);
			return children.some((child, index) => child !== element.children[index])
				? { ...element, children }
				: element;
		}
		if (element.type !== 'zoom') return element;
		const summaryTargets = element.summaryTargets?.map((target) => {
			const targetSlideIndex = resolve(target.targetSlideIndex, target.sectionId);
			return targetSlideIndex === target.targetSlideIndex
				? target
				: { ...target, targetSlideIndex };
		});
		const targetSlideIndex =
			summaryTargets?.[0]?.targetSlideIndex ??
			resolve(element.targetSlideIndex, element.targetSectionId);
		return targetSlideIndex !== element.targetSlideIndex ||
			summaryTargets?.some((target, index) => target !== element.summaryTargets?.[index])
			? { ...element, targetSlideIndex, summaryTargets }
			: element;
	};
	return after.map((slide) => {
		const elements = slide.elements.map(remap);
		return elements.some((element, index) => element !== slide.elements[index])
			? { ...slide, elements }
			: slide;
	});
}
