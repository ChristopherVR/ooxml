import type { DocumentModel, HeaderFooterContent } from '../model';
import type { StoryMargins } from './input';
import type { TextMeasurer } from './measure';
import { headerFooterForPage, pageNumberValues } from './page-fields';
import { flowStory } from './story-flow';
import type { LayoutReviewOptions } from './review-display';
import { twipsToPx } from './units';
import { twips } from '../index';

const DEFAULT_STORY_DISTANCE = twips(720);

export function createStoryMargins(
	model: DocumentModel,
	measurer: TextMeasurer,
	options: LayoutReviewOptions,
	note: (message: string) => void,
): StoryMargins | undefined {
	const { sections, ...base } = model;
	if (!sections?.some((section) => section.headers || section.footers)) return undefined;
	const cache = new WeakMap<HeaderFooterContent, Map<number, number>>();
	const heightOf = (content: HeaderFooterContent, widthPx: number): number => {
		let heights = cache.get(content);
		if (!heights) cache.set(content, (heights = new Map()));
		const cached = heights.get(widthPx);
		if (cached !== undefined) return cached;
		const height = flowStory(base, content.blocks, widthPx, measurer, options, note).heightPx;
		heights.set(widthPx, height);
		return height;
	};
	const usesNumberRestarts = sections.some((section) => section.pageNumbering?.start !== undefined);
	return (pages) => {
		const page = pages.at(-1)!;
		const section = sections[page.sectionIndex];
		const physicalNumber = page.index + 1;
		const number =
			model.evenAndOddHeaders && usesNumberRestarts
				? (pageNumberValues(model, pages).at(-1) ?? physicalNumber)
				: physicalNumber;
		const width = Math.max(1, page.widthPx - page.marginLeftPx - page.marginRightPx);
		const margin = (kind: 'headers' | 'footers', nominal: number, distance: number) => {
			// A negative margin tells Word to use the exact distance, allowing story overlap.
			if (nominal < 0) return -nominal;
			const content = headerFooterForPage(model, page, number, kind);
			const height = content ? heightOf(content, width) : 0;
			return height ? Math.max(nominal, distance + height) : nominal;
		};
		return {
			marginTopPx: margin(
				'headers',
				page.marginTopPx,
				twipsToPx(section?.headerDistanceTwips ?? DEFAULT_STORY_DISTANCE),
			),
			marginBottomPx: margin(
				'footers',
				page.marginBottomPx,
				twipsToPx(section?.footerDistanceTwips ?? DEFAULT_STORY_DISTANCE),
			),
		};
	};
}
