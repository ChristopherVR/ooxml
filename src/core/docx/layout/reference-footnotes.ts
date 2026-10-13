import type { LayoutParagraph } from './input';
import type { LayoutFootnoteBox, LayoutLine } from './result';
import type { TextMeasurer } from './measure';
import { stackParagraphs } from './table-layout';
export type ReferenceFootnotes = (
	paragraph: LayoutParagraph,
	lines: readonly LayoutLine[],
	includeUnanchored: boolean,
) => LayoutFootnoteBox[];
/** Resolve the actual marker fragments, retaining the low-level unanchored-note input contract. */
export function createReferenceFootnotes(
	width: () => number,
	measurer: TextMeasurer,
	report: (message: string) => void,
): ReferenceFootnotes {
	const cache = new Map<string, LayoutFootnoteBox>();
	return (paragraph, lines, includeUnanchored) => {
		const selected = new Set(
			lines.flatMap((line) =>
				line.fragments.map((fragment) => paragraph.runs[fragment.runIndex]?.footnoteReferenceId),
			),
		);
		const anchored = new Set(paragraph.runs.map((run) => run.footnoteReferenceId));
		return (paragraph.footnotes ?? [])
			.filter((note) => selected.has(note.id) || (includeUnanchored && !anchored.has(note.id)))
			.map((note) => {
				const key = `${note.id}:${width()}`;
				const prior = cache.get(key);
				if (prior) return prior;
				const stacked = stackParagraphs(note.paragraphs, width(), measurer, report);
				const box = { id: note.id, yPx: 0, heightPx: stacked.heightPx, paragraphs: stacked.boxes };
				cache.set(key, box);
				return box;
			});
	};
}
