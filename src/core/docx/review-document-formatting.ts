import type { DocumentModel } from './model';
import { mapDocumentParagraphs } from './document-paragraphs';
import {
	reviewParagraphFormatting,
	reviewRunFormatting,
	type ReviewDisplayMode,
} from './review-formatting-display';

export interface ReviewFormattingDiagnostic {
	paragraphId: string;
	runIndex?: number;
	message: string;
}

export function reviewFormattingWarnings(diagnostics: ReviewFormattingDiagnostic[]): string[] {
	return diagnostics.map(
		(diagnostic) =>
			`Original formatting unavailable in paragraph ${diagnostic.paragraphId}${diagnostic.runIndex === undefined ? '' : `, run ${diagnostic.runIndex}`}: ${diagnostic.message}`,
	);
}

/**
 * Project formatting across every story while retaining text and structure. Keeping source
 * offsets unchanged lets read-only layouts map clicks back to the authoritative editor.
 * This does not hide text revisions or merge revised paragraph marks.
 */
export function reviewDocumentFormatting(
	model: DocumentModel,
	mode: ReviewDisplayMode,
): {
	model: DocumentModel;
	diagnostics: ReviewFormattingDiagnostic[];
} {
	const diagnostics: ReviewFormattingDiagnostic[] = [];
	if (mode !== 'original') return { model, diagnostics };
	const projected = mapDocumentParagraphs(model, (paragraph) => {
		const prior = reviewParagraphFormatting(paragraph, mode);
		if (prior.error) diagnostics.push({ paragraphId: paragraph.id, message: prior.error });
		const runs = paragraph.runs.map((run, runIndex) => {
			const priorRun = reviewRunFormatting(run, mode);
			if (priorRun.error)
				diagnostics.push({ paragraphId: paragraph.id, runIndex, message: priorRun.error });
			return priorRun.value;
		});
		return { ...prior.value, runs };
	});
	return { model: projected, diagnostics };
}
