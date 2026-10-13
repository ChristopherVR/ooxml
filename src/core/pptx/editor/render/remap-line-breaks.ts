/**
 * Keep soft line breaks (`a:br`) soft through an inline text edit.
 *
 * Editor surfaces expose a body as plain text, where a paragraph break and an
 * `a:br` are both "\n". `remapTextToSegments` therefore splits the original
 * body into LINES at both kinds of break and pairs each edited line with its
 * source line. Re-joining every pair of lines with a paragraph break turned
 * each `a:br` into a new paragraph on the first commit: the paragraph's lines
 * came back as separate paragraphs without its `a:pPr`, so spacing after
 * (`a:spcAft`) appeared between what were lines of one paragraph.
 *
 * PowerPoint keeps a soft break inside its paragraph when the text around it
 * is edited, so two edited lines whose sources are lines of the SAME original
 * paragraph, in order, are joined by that paragraph's own line break again.
 * A line with no source (typed with Enter) still starts a new paragraph.
 */
import type { TextSegment } from 'ooxml-core/pptx';

import type { RemapSourceParagraph } from './remap-paragraph-sources';

/** Whether a segment is a soft line break rather than a paragraph separator. */
export function isSoftLineBreak(segment: TextSegment): boolean {
	return segment.isLineBreak === true && !segment.isParagraphBreak;
}

/**
 * The soft break to put before edited line `index`, or `undefined` when a
 * paragraph break separates it from the line before. `sources` are the source
 * line indices `alignParagraphSources` paired the edited lines with.
 */
export function softBreakBetween(
	sources: readonly (number | undefined)[],
	index: number,
	lines: readonly RemapSourceParagraph[],
): TextSegment | undefined {
	const previous = index > 0 ? sources[index - 1] : undefined;
	const next = sources[index];
	if (previous === undefined || next === undefined || next <= previous) {
		return undefined;
	}
	for (let line = previous; line < next; line++) {
		const terminator = lines[line]?.terminator;
		if (!terminator || !isSoftLineBreak(terminator)) {
			return undefined;
		}
	}
	const terminator = lines[next - 1].terminator as TextSegment;
	return { ...terminator, text: '\n', style: { ...terminator.style } };
}
