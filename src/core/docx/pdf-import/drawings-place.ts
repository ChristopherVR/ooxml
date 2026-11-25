/**
 * Turns each page's drawing source into drawings, anchors them to the page's paragraphs and notes
 * on the page what was rebuilt and what was left out.
 */
import {
	attachDrawingText,
	MAX_PIECES,
	MAX_SEGMENTS,
	pageDrawings,
	withoutNoteSeparators,
	type DrawingSource,
	type VectorDrawing,
} from './drawings';
import { attachDrawings } from './floats';
import type { FlowItem } from './flow';
import type { Line } from './lines';
import type { SectionLayout } from './package';

const shapes = (count: number) => `${count} shape${count === 1 ? '' : 's'}`;

export interface PreparedPageDrawings {
	drawings: VectorDrawing[];
	pieces: number;
	leftOut: number;
	overLimit: number;
}

/** Builds each page's drawings and moves lines enclosed inside shapes into those shapes. */
export function prepareDrawings(
	pages: { index: number; width: number; height: number; lines: Line[]; drawing: DrawingSource }[],
	notes: readonly { lines: readonly Line[] }[],
	bodySize: number,
): Map<number, PreparedPageDrawings> {
	const noteTops = new Map<number, number>();
	for (const line of notes.flatMap((note) => note.lines))
		noteTops.set(line.page, Math.min(noteTops.get(line.page) ?? Infinity, line.top));
	const prepared = new Map<number, PreparedPageDrawings>();
	for (const page of pages) {
		const source = withoutNoteSeparators(
			page.drawing,
			noteTops.get(page.index),
			bodySize,
			page.width,
		);
		const { drawings, pieces, overLimit } = pageDrawings(page.index, source);
		page.lines = attachDrawingText(drawings, page.lines, page.width * page.height);
		prepared.set(page.index, { drawings, pieces, leftOut: source.leftOut, overLimit });
	}
	return prepared;
}

export function placeDrawings(
	pages: { index: number; width: number; drawing: DrawingSource; notes: string[] }[],
	notes: readonly { lines: readonly Line[] }[],
	blocks: FlowItem[],
	sections: SectionLayout[],
	bodySize: number,
	prepared?: Map<number, PreparedPageDrawings>,
): void {
	const noteTops = new Map<number, number>();
	if (!prepared)
		for (const line of notes.flatMap((note) => note.lines))
			noteTops.set(line.page, Math.min(noteTops.get(line.page) ?? Infinity, line.top));
	for (const page of pages) {
		const entry =
			prepared?.get(page.index) ??
			(() => {
				const source = withoutNoteSeparators(
					page.drawing,
					noteTops.get(page.index),
					bodySize,
					page.width,
				);
				return { ...pageDrawings(page.index, source), leftOut: source.leftOut };
			})();
		attachDrawings(blocks, sections, entry.drawings, bodySize);
		if (entry.pieces) {
			const hasShapeText = entry.drawings.some((d) =>
				d.pieces.some((p) => (p.textLines?.length ?? 0) > 0),
			);
			page.notes.push(
				hasShapeText
					? `Vector drawings (${shapes(entry.pieces)}) became Word drawings placed where the PDF drew them.`
					: `Vector drawings (${shapes(entry.pieces)}) became Word drawings placed where the PDF drew them, behind the text; text in them stays ordinary text.`,
			);
		}
		if (entry.leftOut)
			page.notes.push(
				`Vector drawings (${shapes(entry.leftOut)}) were left out: they use unsupported pattern fills or non-rectangular clipping paths.`,
			);
		if (entry.overLimit)
			page.notes.push(
				`Vector drawings (${shapes(entry.overLimit)}) were left out: a page converts at most ${MAX_PIECES} shapes (${MAX_SEGMENTS} path segments).`,
			);
	}
}
