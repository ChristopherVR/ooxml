// Floating objects (`wp:anchor`) anchored in a table cell, laid out in that cell as Word does:
// measured with Word 16 in `docx/__fixtures__/cell-floats` (compatibility modes 14 and 15).
// - From mode 15 every such object is laid out in its cell (`layoutInCell` is ignored); before
//   mode 15 only with `layoutInCell="1"`, and `"0"` positions it on the page (`position-floats.ts`).
// - Its frames are the cell's: `column`, `margin` and `character` the cell's text area, `page` the
//   cell's edges, `leftMargin`/`rightMargin` (and inside/outside) the cell margins beside the text,
//   and every vertical frame but `paragraph` and `line` the cell from its top. It never starts
//   above the cell's top, and a wrapping object stays between the cell's edges.
// - Wrapping (square, tight, through, top and bottom) takes room from the cell's text the way it
//   does in the body, and the row grows to hold the object; in front of or behind text it does
//   neither.
import { anchorPoint, placeInFrames, type FloatFrames, type Span } from './floats';
import type { LayoutFloat, LayoutParagraph } from './input';
import type { TextMeasurer } from './measure';
import { layoutParagraph, type ParagraphLayoutResult } from './paragraph-layout';
import type { LayoutCellFloat, LayoutParagraphBox } from './result';
import { offsetParagraphFrame } from './result';
import { floatExclusion, lineBoxesFor, wrapsText, type Exclusion } from './wrap';

/** The cell margins beside a cell's text area, for the frames that reach past it. */
export interface CellSides {
	paddingLeftPx: number;
	paddingRightPx: number;
}

/** Whether a paragraph's float is laid out in its table cell rather than on the page. */
export function laidOutInCell(float: LayoutFloat): boolean {
	return float.layoutInCell !== false;
}

/**
 * The frames of an object in a cell, with x from the cell's text area and y from the cell's top:
 * `widthPx` is the text area's width and `heightPx` the height of the cell's text.
 */
export function cellFrames(
	cell: CellSides & { widthPx: number; heightPx: number },
	paragraph: Span,
	anchor?: { line: Span; characterPx: number },
): FloatFrames {
	const text = { start: 0, size: cell.widthPx };
	const box = { start: 0, size: cell.heightPx };
	return {
		page: {
			start: -cell.paddingLeftPx,
			size: cell.paddingLeftPx + cell.widthPx + cell.paddingRightPx,
		},
		margin: text,
		column: text,
		leftMargin: { start: -cell.paddingLeftPx, size: cell.paddingLeftPx },
		rightMargin: { start: cell.widthPx, size: cell.paddingRightPx },
		pageV: box,
		marginV: box,
		topMargin: box,
		bottomMargin: box,
		paragraph,
		...(anchor
			? {
					line: { start: paragraph.start + anchor.line.start, size: anchor.line.size },
					characterPx: anchor.characterPx,
				}
			: {}),
	};
}

/** Where a float sits in its cell, kept below the cell's top and, when it wraps, between its edges. */
export function placeInCell(
	float: LayoutFloat,
	cell: CellSides & { widthPx: number; heightPx: number },
	paragraph: Span,
	lines: ParagraphLayoutResult['lines'],
): { xPx: number; yPx: number; originXPx: number; originYPx: number } {
	const placed = placeInFrames(
		float,
		cellFrames(cell, paragraph, anchorPoint(lines, float.sourceOffset)),
	);
	const right = cell.widthPx + cell.paddingRightPx - float.widthPx;
	return {
		...placed,
		xPx: wrapsText(float) ? Math.max(-cell.paddingLeftPx, Math.min(placed.xPx, right)) : placed.xPx,
		yPx: Math.max(0, placed.yPx),
	};
}

const overlaps = (exclusions: readonly Exclusion[], topPx: number, heightPx: number) =>
	exclusions.some((item) => item.yPx < topPx + heightPx && item.yPx + item.heightPx > topPx);

/**
 * Stacks a cell's paragraphs (as `stackParagraphs` does) with the floating objects anchored in
 * them laid out in the cell: each object's position goes on its paragraph's box, the text after
 * and beside a wrapping object runs around it, and the content reaches at least its bottom.
 */
export function stackFloatingCell(
	paragraphs: LayoutParagraph[],
	widthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	topPx: number,
	startTokens: ReadonlyMap<string, import('./text-breaks').TokenPosition> | undefined,
	cell: CellSides,
): { heightPx: number; boxes: LayoutParagraphBox[] } {
	const start = (paragraph: LayoutParagraph) => startTokens?.get(paragraph.id) ?? 0;
	const plain = paragraphs.map((paragraph) =>
		layoutParagraph(paragraph, widthPx, measurer, note, undefined, start(paragraph)),
	);
	const textBottom = plain.reduce(
		(y, layout) => y + layout.spacingBeforePx + layout.contentHeightPx + layout.spacingAfterPx,
		topPx,
	);
	const frame = { ...cell, widthPx, heightPx: textBottom };
	const exclusions: Exclusion[] = [];
	const boxes: LayoutParagraphBox[] = [];
	let y = topPx;
	let bottom = 0;
	paragraphs.forEach((paragraph, index) => {
		let layout = plain[index]!;
		y += layout.spacingBeforePx;
		const top = y;
		const wrapped = () =>
			layoutParagraph(
				paragraph,
				widthPx,
				measurer,
				note,
				lineBoxesFor(exclusions, top, 0, widthPx),
				start(paragraph),
			);
		if (overlaps(exclusions, top, layout.contentHeightPx)) layout = wrapped();
		const floats: LayoutCellFloat[] = [];
		// A paragraph continued from an earlier page keeps its objects where it started.
		const own = start(paragraph) ? [] : (paragraph.floats ?? []).filter(laidOutInCell);
		let wraps = false;
		for (const [position, float] of own.entries()) {
			const placed = placeInCell(
				float,
				frame,
				{ start: top, size: layout.contentHeightPx },
				layout.lines,
			);
			floats.push({
				index: float.index ?? position,
				xPx: placed.xPx,
				yPx: placed.yPx - top,
				originXPx: placed.originXPx,
				originYPx: placed.originYPx - top,
			});
			if (!wrapsText(float)) continue;
			wraps = true;
			exclusions.push(floatExclusion({ ...float, xPx: placed.xPx, yPx: placed.yPx }));
			bottom = Math.max(bottom, placed.yPx + float.heightPx);
		}
		if (wraps && overlaps(exclusions, top, layout.contentHeightPx)) layout = wrapped();
		const paragraphFrame = offsetParagraphFrame(layout.frame, layout.lines[0]?.gapBeforePx);
		boxes.push({
			kind: 'paragraph',
			blockId: paragraph.id,
			yPx: top,
			heightPx: layout.contentHeightPx,
			lines: layout.lines.map((line) => ({ ...line })),
			...(paragraphFrame ? { frame: paragraphFrame } : {}),
			...(floats.length ? { floats } : {}),
		});
		y += layout.contentHeightPx + layout.spacingAfterPx;
	});
	return { heightPx: Math.max(y, bottom), boxes };
}

/** Whether a cell's paragraphs hold an object laid out in the cell. */
export function hasCellFloats(paragraphs: readonly LayoutParagraph[]): boolean {
	return paragraphs.some((paragraph) => paragraph.floats?.some(laidOutInCell));
}
