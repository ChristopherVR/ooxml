import type { Paragraph, TextRun } from './model';
import { restoreParagraphFormatting } from './restore-paragraph-format';
import { restoreRunFormatting } from './restore-run-format';

export type ReviewDisplayMode = 'all' | 'simple' | 'final' | 'original';

export interface ReviewFormattingProjection<T> {
	value: T;
	/** A missing or unsupported snapshot leaves the current formatting visible. */
	error?: string;
}

function project<T>(value: T, restore: (copy: T) => void): ReviewFormattingProjection<T> {
	const copy = { ...value };
	try {
		restore(copy);
		return { value: copy };
	} catch (error) {
		return { value, error: error instanceof Error ? error.message : String(error) };
	}
}

/** Prior formatting for display only: text, anchors and the source's revisions stay intact. */
export function reviewParagraphFormatting(
	paragraph: Paragraph,
	mode: ReviewDisplayMode,
): ReviewFormattingProjection<Paragraph> {
	return mode === 'original' && paragraph.formatRevision?.kind === 'paragraphChange'
		? project(paragraph, restoreParagraphFormatting)
		: { value: paragraph };
}

/** Does not accept or reject text revisions, including those overlapping formatting history. */
export function reviewRunFormatting(
	run: TextRun,
	mode: ReviewDisplayMode,
): ReviewFormattingProjection<TextRun> {
	return mode === 'original' &&
		(run.formatRevision?.kind === 'formatChange' || run.revision?.kind === 'formatChange')
		? project(run, restoreRunFormatting)
		: { value: run };
}
