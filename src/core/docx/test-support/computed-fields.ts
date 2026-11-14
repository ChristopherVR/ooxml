// Test helper: the body paragraphs of a document as the layout draws them.
import type { DocumentModel } from '../model';
import { adaptDocumentModel } from '../layout/adapter';
import { withComputedFieldRuns } from '../layout/computed-field-runs';
import type { LayoutParagraph } from '../layout/input';

/**
 * Each body paragraph as the layout draws it (its list label included) and the source length of
 * its own runs.
 */
export function drawn(model: DocumentModel) {
	const input = withComputedFieldRuns(adaptDocumentModel(model), model);
	return input.sections
		.flatMap((section) => section.blocks)
		.filter((block): block is LayoutParagraph => block.kind === 'paragraph')
		.map((paragraph) => {
			const shown = paragraph.runs.filter((run) => !run.hidden);
			const runs = shown.filter((run) => !run.synthetic);
			return {
				text: shown.map((run) => run.text).join(''),
				length: runs.reduce((sum, run) => sum + (run.sourceLength ?? run.text.length), 0),
			};
		});
}
