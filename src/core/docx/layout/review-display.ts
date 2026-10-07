import type { DocumentModel, Paragraph, TextRun } from '../model';
import { mapDocumentParagraphs } from '../document-paragraphs';
import { isRunHiddenForReview, type ReviewDisplayMode } from '../review-formatting-display';

export interface LayoutReviewOptions {
	reviewDisplayMode?: ReviewDisplayMode;
}

/** Inline atoms occupy one source position even when their display text is empty. */
export function reviewSourceLength(run: TextRun): number {
	return run.equation ||
		run.break ||
		run.fieldChar ||
		run.fieldCode !== undefined ||
		run.noteReference ||
		run.image
		? 1
		: run.text.length;
}

/** Visible references and objects for derived numbering, notes and float placement only. */
export function visibleReviewParagraph(paragraph: Paragraph, mode: ReviewDisplayMode): Paragraph {
	return mode === 'all'
		? paragraph
		: {
				...paragraph,
				runs: paragraph.runs.filter((run) => !isRunHiddenForReview(run, mode)),
			};
}

export function visibleReviewModel(model: DocumentModel, mode: ReviewDisplayMode): DocumentModel {
	return mode === 'all'
		? model
		: mapDocumentParagraphs(model, (paragraph) => visibleReviewParagraph(paragraph, mode));
}
