import { documentParagraphs } from '../cell-blocks';
import type { DocumentModel, HeaderFooterContent } from '../model';
import type { LayoutDocumentInput } from './input';
import type { TextMeasurer } from './measure';
import { pageNumberValues } from './page-fields';
import type { LayoutPageBox } from './result';
import type { LayoutReviewOptions } from './review-display';
import { createPageStories } from './story-flow';
import { floatExclusions, wrapsText, type Exclusion } from './wrap';

export function createStoryWrap(
	model: DocumentModel,
	measurer: TextMeasurer,
	options: LayoutReviewOptions,
	note: (message: string) => void,
): LayoutDocumentInput['storyWrap'] {
	const canWrap = (content: HeaderFooterContent) =>
		documentParagraphs(content.blocks).some((p) =>
			p.runs.some(({ image }) => image?.anchored && wrapsText(image.placement ?? {})),
		);
	const content = new Set(
		model.sections?.flatMap((section) => [
			...Object.values(section.headers ?? {}),
			...Object.values(section.footers ?? {}),
		]),
	);
	if (![...content].some(canWrap)) return undefined;

	const restarts = model.sections?.some((section) => section.pageNumbering?.start !== undefined);
	return (previousPages) => {
		const stories = createPageStories(model, previousPages, measurer, options, note);
		const previousNumbers = restarts ? pageNumberValues(model, previousPages) : undefined;
		const cache = new WeakMap<LayoutPageBox, Exclusion[]>();
		return (pages, page) => {
			const stored = cache.get(page);
			if (stored) return stored;
			const numbers =
				pages === previousPages
					? previousNumbers
					: restarts
						? pageNumberValues(model, pages)
						: undefined;
			const number = numbers?.[page.index] ?? page.index + 1;
			const exclusions = stories(page.index, page, number).flatMap(({ floats }) =>
				floatExclusions(floats),
			);
			cache.set(page, exclusions);
			return exclusions;
		};
	};
}
