import { cellParagraphs } from './table-cell-layout';
import { anchorPoint, floatBox, pageFrames, placeInFrames } from './floats';
import type { LayoutDocumentInput, LayoutFloat, LayoutParagraph } from './input';
import {
	cellVerticalOffset,
	type LayoutColumnBox,
	type LayoutFloatBox,
	type LayoutPageBox,
	type LayoutTableBox,
} from './result';

/** Every paragraph with floating objects, in the body and in table cells (nested ones too). */
function anchorParagraphs(input: LayoutDocumentInput): Map<string, LayoutFloat[]> {
	const floatsById = new Map<string, LayoutFloat[]>();
	const add = (paragraph: LayoutParagraph) => {
		if (paragraph.floats?.length) floatsById.set(paragraph.id, paragraph.floats);
	};
	for (const section of input.sections)
		for (const block of section.blocks)
			if (block.kind === 'paragraph') add(block);
			else
				for (const row of block.rows)
					for (const cell of row.cells) cellParagraphs(cell).forEach(add);
	return floatsById;
}

/**
 * A table's floating objects on the page: one laid out in its cell (`cell-floats.ts`) is moved
 * from its paragraph box onto the page; one positioned on the page (`layoutInCell="0"` before
 * mode 15) takes the page's frames, with the row as its paragraph, as Word 2010 places it.
 */
function tableFloats(
	page: LayoutPageBox,
	originYPx: number,
	column: LayoutColumnBox,
	table: LayoutTableBox,
	floatsById: ReadonlyMap<string, LayoutFloat[]>,
	placed: Set<string>,
	boxes: LayoutFloatBox[],
): void {
	const left = page.marginLeftPx + column.xPx + (table.xPx ?? 0);
	const top = originYPx + table.yPx;
	for (const row of table.rows)
		row.cells.forEach((paragraphs, cellIndex) => {
			const geometry = row.geometry?.[cellIndex];
			const contentLeft = left + (geometry ? geometry.xPx + geometry.paddingLeftPx : 0);
			const rowTop = top + row.yPx;
			const shift = geometry ? cellVerticalOffset(geometry, row.heightPx) : 0;
			for (const box of paragraphs) {
				const floats = floatsById.get(box.blockId);
				if (!floats || placed.has(box.blockId)) continue;
				placed.add(box.blockId);
				const boxLeft = contentLeft + (box.xPx ?? 0);
				const boxTop = rowTop + shift + box.yPx;
				for (const [position, float] of floats.entries()) {
					const index = float.index ?? position;
					const local = box.floats?.find((item) => item.index === index);
					if (local)
						boxes.push(
							floatBox(
								box.blockId,
								float,
								index,
								{
									xPx: boxLeft + local.xPx,
									yPx: boxTop + local.yPx,
									originXPx: boxLeft + local.originXPx,
									originYPx: boxTop + local.originYPx,
								},
								true,
							),
						);
					else if (float.layoutInCell === false)
						boxes.push(
							floatBox(
								box.blockId,
								float,
								index,
								placeInFrames(
									float,
									pageFrames(page, column, { topPx: rowTop, heightPx: row.heightPx }),
								),
							),
						);
				}
			}
		});
}

/**
 * Positions each paragraph's floating pictures on the page where the paragraph starts, from their
 * `wp:positionH`/`wp:positionV` reference frame, alignment and offset, and the floating objects
 * of table cells where their cells placed them. Returns whether any were placed.
 */
export function positionFloats(
	input: LayoutDocumentInput,
	pages: LayoutPageBox[],
	/** A header or footer starts at its own page distance. */
	storyTopPx?: number,
): boolean {
	const floatsById = anchorParagraphs(input);
	if (!floatsById.size) return false;
	const placed = new Set<string>();
	for (const page of pages) {
		const boxes: LayoutFloatBox[] = [];
		const originYPx = storyTopPx ?? page.marginTopPx;
		for (const column of page.columns)
			for (const block of column.blocks) {
				if (block.kind === 'table') {
					tableFloats(page, originYPx, column, block, floatsById, placed, boxes);
					continue;
				}
				const floats = floatsById.get(block.blockId);
				if (!floats || placed.has(block.blockId)) continue;
				placed.add(block.blockId);
				const paragraph = { topPx: originYPx + block.yPx, heightPx: block.heightPx };
				for (const [position, float] of floats.entries()) {
					const frames = pageFrames(
						page,
						column,
						paragraph,
						anchorPoint(block.lines, float.sourceOffset),
					);
					boxes.push(
						floatBox(block.blockId, float, float.index ?? position, placeInFrames(float, frames)),
					);
				}
			}
		if (boxes.length)
			page.floats = boxes.sort((a, b) => (a.relativeHeight ?? 0) - (b.relativeHeight ?? 0));
	}
	return true;
}
