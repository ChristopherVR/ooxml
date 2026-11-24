/**
 * Turns each page's drawing source into drawings, anchors them to the page's paragraphs and notes
 * on the page what was rebuilt and what was left out.
 */
import {
	MAX_PIECES,
	MAX_SEGMENTS,
	pageDrawings,
	withoutNoteSeparators,
	type DrawingSource,
} from './drawings';
import { attachDrawings } from './floats';
import type { FlowItem } from './flow';
import type { Line } from './lines';
import type { SectionLayout } from './package';

const shapes = (count: number) => `${count} shape${count === 1 ? '' : 's'}`;

export function placeDrawings(
	pages: { index: number; width: number; drawing: DrawingSource; notes: string[] }[],
	notes: readonly { lines: readonly Line[] }[],
	blocks: FlowItem[],
	sections: SectionLayout[],
	bodySize: number,
): void {
	const noteTops = new Map<number, number>();
	for (const line of notes.flatMap((note) => note.lines))
		noteTops.set(line.page, Math.min(noteTops.get(line.page) ?? Infinity, line.top));
	for (const page of pages) {
		const source = withoutNoteSeparators(
			page.drawing,
			noteTops.get(page.index),
			bodySize,
			page.width,
		);
		const { drawings, pieces, overLimit } = pageDrawings(page.index, source);
		attachDrawings(blocks, sections, drawings, bodySize);
		if (pieces)
			page.notes.push(
				`Vector drawings (${shapes(pieces)}) became Word drawings placed where the PDF drew them, behind the text; text in them stays ordinary text.`,
			);
		if (source.leftOut)
			page.notes.push(
				`Vector drawings (${shapes(source.leftOut)}) were left out: they use pattern fills or a clipping path cuts them.`,
			);
		if (overLimit)
			page.notes.push(
				`Vector drawings (${shapes(overLimit)}) were left out: a page converts at most ${MAX_PIECES} shapes (${MAX_SEGMENTS} path segments).`,
			);
	}
}
