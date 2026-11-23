import type { Block, DocumentModel } from '../index';
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
import { createTextBoxLayouts, type TextBoxLayouts } from './text-box-layout';
import { createStoryMargins } from './story-margins';

/** Paginates an engine-native `LayoutDocumentInput` (bypasses the DocumentModel adapter). */
export function layoutDocument(input: LayoutDocumentInput, measurer: TextMeasurer): LayoutResult {
	let result = layoutSections(input, measurer);
	const pictures = positionFloats(input, result.pages);
	const tables = input.sections.some((section) =>
		section.blocks.some((block) => block.kind === 'table' && block.floating),
	);
	if (!pictures && !tables) return result;
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
	const result = layout(input, measurer);
	return {
		pages: result.pages,
		approximations: [...new Set([...approximations, ...result.approximations])],
	};
}

/**
 * The text boxes of one of `model`'s stories (its body, or a header or footer's blocks) laid out by
 * the page engine with the document's styles, linked chains included.
 */
export function storyTextBoxLayouts(
	model: DocumentModel,
	story: readonly Block[],
	measurer: TextMeasurer,
	options: LayoutReviewOptions = {},
	note: (message: string) => void = () => {},
): TextBoxLayouts {
	// The box's blocks adapt as a document of their own: the styles, theme and numbering are the
	// document's; its sections, notes and separators are not the box's.
	const {
		sections: _sections,
		footnoteSeparators: _separators,
		footnotes: _footnotes,
		endnotes: _endnotes,
		...base
	} = model;
	return createTextBoxLayouts(
		story,
		(blocks) => adaptDocumentModel({ ...base, blocks }, note, options).sections[0]?.blocks ?? [],
		measurer,
		note,
	);
}
