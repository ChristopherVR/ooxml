import type { PptxElement, PptxSlide } from '../../index';

/** Shared chart parts cannot be edited independently through separate shape models. */
export function chartSourceReferenceCounts(slides: readonly PptxSlide[]): Map<string, number> {
	const counts = new Map<string, number>();
	function visit(elements: readonly PptxElement[]): void {
		for (const element of elements) {
			if (element.type === 'group') visit(element.children);
			else if (element.type === 'chart' && element.chartData?.chartPartPath) {
				const part = element.chartData.chartPartPath;
				counts.set(part, (counts.get(part) ?? 0) + 1);
			}
		}
	}
	for (const slide of slides) visit(slide.elements);
	return counts;
}

/** Resolve the actual model source; never trust an editable flag carried by a caller's match. */
export function findMatchChartSource(
	slides: readonly PptxSlide[],
	match: { slideIndex: number; elementId: string; elementPath?: number[] },
): string | undefined {
	let element = slides[match.slideIndex]?.elements.find((item) => item.id === match.elementId);
	for (const index of match.elementPath ?? []) {
		element = element?.type === 'group' ? element.children[index] : undefined;
	}
	return element?.type === 'chart' ? element.chartData?.chartPartPath : undefined;
}
