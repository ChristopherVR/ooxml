import type { DocumentModel } from '../index';
import { adaptDocumentModel } from './adapter';
import { withComputedFieldRuns } from './computed-field-runs';
import { layoutSections } from './page-flow';
import { FLOAT_WRAP_NOTE } from './floats';
import { positionFloats } from './position-floats';
import { sameExclusions, wrapExclusions } from './wrap';
import type { LayoutDocumentInput } from './input';
import type { LayoutResult } from './result';
import type { TextMeasurer } from './measure';
import type { LayoutReviewOptions } from './review-display';
import { reviewDocumentFormatting, reviewFormattingWarnings } from '../review-document-formatting';
import { storyTextBoxLayouts } from './story-flow';
import { createStoryMargins } from './story-margins';
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
	if (!pictures && !tables && !frames) return result;
	// Wrapped pictures and floating tables take room from the text, which can move their anchors:
	// lay out again around them until their positions settle (usually one extra pass).
	let exclusions = wrapExclusions(result.pages);
	for (let pass = 0; exclusions.size && pass < 3; pass++) {
		result = layoutSections(input, measurer, exclusions);
		positionFloats(input, result.pages);
		const next = wrapExclusions(result.pages);
		if (sameExclusions(next, exclusions)) break;
		exclusions = next;
	}
	const notes = pictures ? [...result.approximations, FLOAT_WRAP_NOTE] : result.approximations;
	return { ...result, approximations: [...new Set(notes)] };
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
	if (options.oversizedCantSplitRow) input.oversizedCantSplitRow = options.oversizedCantSplitRow;
	const result = layout(input, measurer);
	return {
		pages: result.pages,
		approximations: [...new Set([...approximations, ...result.approximations])],
	};
}
