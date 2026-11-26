import type { DocumentModel } from '../index';
import { adaptDocumentModel } from './adapter';
import { withComputedFieldRuns } from './computed-field-runs';
import { layoutSections } from './page-flow';
import { FLOAT_WRAP_NOTE } from './floats';
import { positionFloats } from './position-floats';
import { sameExclusions, wrapExclusions, type Exclusion } from './wrap';
import type { LayoutDocumentInput, StoryExclusions } from './input';
import type { LayoutPageBox, LayoutResult } from './result';
import type { TextMeasurer } from './measure';
import type { LayoutReviewOptions } from './review-display';
import { reviewDocumentFormatting, reviewFormattingWarnings } from '../review-document-formatting';
import { storyTextBoxLayouts } from './story-flow';
import { createStoryMargins } from './story-margins';
import { createStoryWrap } from './story-floats';
export { storyTextBoxLayouts } from './story-flow';

/** Paginates an engine-native `LayoutDocumentInput` (bypasses the DocumentModel adapter). */
export function layoutDocument(input: LayoutDocumentInput, measurer: TextMeasurer): LayoutResult {
	let result = layoutSections(input, measurer);
	const pictures = positionFloats(input, result.pages);
	const tables = input.sections.some((section) =>
		section.blocks.some((block) => block.kind === 'table' && block.floating),
	);
	const frames = input.sections.some((section) =>
		section.blocks.some((block) => block.kind === 'paragraph' && block.textFrame),
	);
	if (!pictures && !tables && !frames && !input.storyWrap) return result;
	// Wrapped pictures and floating tables take room from the text, which can move their anchors:
	// lay out again around them until their positions settle (usually one extra pass).
	let body = wrapExclusions(result.pages);
	let story = input.storyWrap?.(result.pages);
	let exclusions = combinedExclusions(result.pages, body, story);
	for (let pass = 0; exclusions.size && pass < 3; pass++) {
		result = layoutSections(input, measurer, body, story);
		positionFloats(input, result.pages);
		body = wrapExclusions(result.pages);
		story = input.storyWrap?.(result.pages);
		const next = combinedExclusions(result.pages, body, story);
		if (sameExclusions(next, exclusions)) break;
		exclusions = next;
	}
	const notes =
		pictures || result.pages.some((page) => story?.(result.pages, page).length)
			? [...result.approximations, FLOAT_WRAP_NOTE]
			: result.approximations;
	return { ...result, approximations: [...new Set(notes)] };
}

function combinedExclusions(
	pages: LayoutPageBox[],
	body: Map<number, Exclusion[]>,
	story: StoryExclusions | undefined,
): Map<number, Exclusion[]> {
	if (!story) return body;
	const combined = new Map<number, Exclusion[]>();
	for (const page of pages) {
		const list = [...(body.get(page.index) ?? []), ...story(pages, page)];
		if (list.length) combined.set(page.index, list);
	}
	return combined;
}

/** Paginates a `docx-core` `DocumentModel` directly: `adaptDocumentModel` + `layoutDocument`. */
export function layoutDocumentModel(
	model: DocumentModel,
	measurer: TextMeasurer,
	options: LayoutReviewOptions = {},
	layout: (input: LayoutDocumentInput, measurer: TextMeasurer) => LayoutResult = layoutDocument,
): LayoutResult {
	const projected = reviewDocumentFormatting(model, options.reviewDisplayMode ?? 'all');
	const approximations = reviewFormattingWarnings(projected.diagnostics);
	// Computed fields (DOCPROPERTY, AUTONUM, LISTNUM, SYMBOL...) show what the document gives them.
	const note = (message: string) => approximations.push(message);
	const textBoxes = storyTextBoxLayouts(
		projected.model,
		projected.model.blocks,
		measurer,
		options,
		note,
	);
	const input = withComputedFieldRuns(
		adaptDocumentModel(projected.model, note, options, textBoxes),
		projected.model,
	);
	const storyMargins = createStoryMargins(projected.model, measurer, options, note);
	if (storyMargins) input.storyMargins = storyMargins;
	const storyWrap = createStoryWrap(projected.model, measurer, options, note);
	if (storyWrap) input.storyWrap = storyWrap;
	if (options.oversizedCantSplitRow) input.oversizedCantSplitRow = options.oversizedCantSplitRow;
	const result = layout(input, measurer);
	return {
		pages: result.pages,
		approximations: [...new Set([...approximations, ...result.approximations])],
	};
}
