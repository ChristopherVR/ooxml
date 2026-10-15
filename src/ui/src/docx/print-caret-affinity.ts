/**
 * Which side of a boundary the caret sits on when one text offset belongs to two rendered lines
 * (the end of a wrapped line and the start of the next, across a column or a page). Upstream is the
 * end of the earlier line; downstream is the start of the later one.
 */
export type CaretAffinity = 'upstream' | 'downstream';

const start = (line: HTMLElement) => Number(line.dataset.sourceStart);
const end = (line: HTMLElement) => Number(line.dataset.sourceEnd);

/** Lines of a block whose source range holds the offset, in reading order. */
export function candidateLines(block: readonly HTMLElement[], offset: number): HTMLElement[] {
	return block.filter((line) => start(line) <= offset && end(line) >= offset);
}

/** Pick the line a caret at `offset` is drawn on. */
export function lineForCaret(
	block: readonly HTMLElement[],
	offset: number,
	affinity: CaretAffinity,
): HTMLElement | undefined {
	const candidates = candidateLines(block, offset);
	if (!candidates.length) {
		// The layout is a moment behind the text (relayout is debounced): the last line starting at
		// or before the offset is where the new text will appear.
		return block.filter((line) => start(line) <= offset).at(-1);
	}
	return affinity === 'upstream' ? candidates[0] : candidates.at(-1);
}

/**
 * Affinity a caret placed at `offset` on `line` has: upstream when the offset is the line's end and
 * a later line of the same block starts there, downstream otherwise.
 */
export function affinityOnLine(
	line: HTMLElement,
	offset: number,
	block: readonly HTMLElement[],
): CaretAffinity {
	if (offset < end(line) || offset <= start(line)) return 'downstream';
	const position = block.indexOf(line);
	return position >= 0 && block.slice(position + 1).some((later) => start(later) <= offset)
		? 'upstream'
		: 'downstream';
}
