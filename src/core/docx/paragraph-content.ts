import type { Paragraph } from './model';

/** True when the paragraph has only a mark, possibly with formatting on empty runs. */
export function isEmptyParagraph(paragraph: Paragraph): boolean {
	return (
		!paragraph.contentControl &&
		paragraph.runs.every(
			(run) =>
				!run.text &&
				!run.image &&
				!run.equation &&
				!run.field &&
				!run.fieldChar &&
				run.fieldCode === undefined &&
				!run.formField &&
				!run.break &&
				!run.noteReference &&
				!run.noteMark &&
				!run.noteSeparator &&
				!run.contentControl,
		)
	);
}
