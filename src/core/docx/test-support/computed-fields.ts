// Test helper: the body paragraphs of a document as the layout draws them.
import type { DocumentModel } from '../model';
import { adaptDocumentModel } from '../layout/adapter';
import { withComputedFieldRuns } from '../layout/computed-field-runs';
import type { LayoutParagraph } from '../layout/input';

/** Each body paragraph as the layout draws it, and the source length of its runs. */
export function drawn(model: DocumentModel) {
	const input = withComputedFieldRuns(adaptDocumentModel(model), model);
	return input.sections
		.flatMap((section) => section.blocks)
		.filter((block): block is LayoutParagraph => block.kind === 'paragraph')
		.map((paragraph) => {
			const runs = paragraph.runs.filter((run) => !run.synthetic && !run.hidden);
			return {
				text: runs.map((run) => run.text).join(''),
				length: runs.reduce((sum, run) => sum + (run.sourceLength ?? run.text.length), 0),
			};
		});
}
