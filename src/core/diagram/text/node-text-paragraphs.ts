// Plain-text views of a SmartArt node's typed paragraphs. Moved from
// `pptx/core/core/runtime/smartart-text-paragraphs.ts` (the object-tree parser and builder stay
// there); generic over the reader's run style.
import type { DiagramNodeTextParagraph, DiagramNodeTextRun } from '../model/node';

/** Flattens typed paragraphs into the node text value: `\n` between paragraphs and for breaks, `\t` for tabs. */
export function diagramParagraphsText<S>(paragraphs: DiagramNodeTextParagraph<S>[]): string {
	return paragraphs
		.map((paragraph) =>
			paragraph.items
				.map((item) => {
					if (item.kind === 'run') {
						return item.run.text;
					}
					if (item.kind === 'field') {
						return item.text;
					}
					if (item.kind === 'break') {
						return '\n';
					}
					return item.kind === 'tab' ? '\t' : '';
				})
				.join(''),
		)
		.join('\n');
}

/** The regular runs of the first paragraph, for the legacy `node.runs` API. */
export function firstDiagramParagraphRuns<S>(
	paragraphs: DiagramNodeTextParagraph<S>[] | undefined,
): DiagramNodeTextRun<S>[] | undefined {
	const runs = paragraphs?.[0]?.items.flatMap((item) => (item.kind === 'run' ? [item.run] : []));
	return runs && runs.length > 0 ? runs : undefined;
}
