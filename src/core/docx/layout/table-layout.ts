import type { TextMeasurer } from './measure';
import { layoutParagraph } from './paragraph-layout';
import type { LayoutParagraph, LayoutTableRow } from './input';
import type { LayoutCellGeometry, LayoutParagraphBox } from './result';
import { expectDefined } from '../expect-defined';
import { stackCellBlocks } from './table-cell-layout';

export interface RowLayout {
	heightPx: number;
	cells: LayoutParagraphBox[][];
	geometry: LayoutCellGeometry[];
	/** Vertical cell margins apply to every physical row fragment. */
	cellPadding?: { top: number; bottom: number }[];
}

/** Lays out paragraphs stacked top to bottom (a table cell's or a footnote's content). */
export function stackParagraphs(
	paragraphs: LayoutParagraph[],
	widthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	topPx = 0,
	startTokens?: ReadonlyMap<string, number>,
): { heightPx: number; boxes: LayoutParagraphBox[] } {
	let y = topPx;
	const boxes: LayoutParagraphBox[] = [];
	for (const paragraph of paragraphs) {
		const layout = layoutParagraph(
			paragraph,
			widthPx,
			measurer,
			note,
			undefined,
			startTokens?.get(paragraph.id) ?? 0,
		);
		y += layout.spacingBeforePx;
		const top = y;
		const lines = layout.lines.map((line) => ({ ...line }));
		boxes.push({
			kind: 'paragraph',
			blockId: paragraph.id,
			yPx: top,
			heightPx: layout.contentHeightPx,
			lines,
			...(layout.frame ? { frame: layout.frame } : {}),
		});
		y += layout.contentHeightPx + layout.spacingAfterPx;
	}
	return { heightPx: y, boxes };
}

/**
 * Lays out one table row. Cells take their grid position and width when the table has a grid,
 * otherwise the table width split evenly, and their content is inset by the cell margins.
 */
export function layoutRow(
	row: LayoutTableRow,
	tableWidthPx: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	remaining?: LayoutParagraphBox[][],
	startTokens?: ReadonlyMap<string, number>,
): RowLayout {
	const count = row.cells.length || 1;
	let x = 0;
	const geometry: LayoutCellGeometry[] = [];
	const laidOut = row.cells.map((cell) => {
		const continuation = remaining?.[geometry.length];
		const starts = continuation
			? new Map(continuation.map((box) => [box.blockId, box.lines[0]?.firstToken ?? 0]))
			: startTokens;
		const widthPx = cell.widthPx ?? tableWidthPx / count;
		const padding = cell.padding ?? { top: 0, right: 0, bottom: 0, left: 0 };
		const content = stackCellBlocks(
			cell,
			Math.max(1, widthPx - padding.left - padding.right),
			measurer,
			note,
			padding.top,
			starts,
		);
		const heightPx = content.heightPx + padding.bottom;
		geometry.push({
			xPx: cell.xPx ?? x,
			widthPx,
			paddingLeftPx: padding.left,
			paddingRightPx: padding.right,
			contentHeightPx: heightPx,
			...(cell.borders ? { borders: cell.borders } : {}),
			...(cell.shading ? { shading: cell.shading } : {}),
			...(cell.verticalAlign ? { verticalAlign: cell.verticalAlign } : {}),
		});
		x = (cell.xPx ?? x) + widthPx;
		return { heightPx, boxes: content.boxes };
	});
	const content = Math.max(0, ...laidOut.map((cell) => cell.heightPx));
	const heightPx =
		row.heightPx === undefined
			? content
			: row.heightRule === 'exact'
				? row.heightPx
				: Math.max(content, row.heightPx);
	return {
		heightPx,
		cells: laidOut.map((cell) => cell.boxes),
		geometry,
		cellPadding: row.cells.map((cell) => ({
			top: cell.padding?.top ?? 0,
			bottom: cell.padding?.bottom ?? 0,
		})),
	};
}

interface CellSplit {
	before: LayoutParagraphBox[];
	after: LayoutParagraphBox[];
	beforeHeightPx: number;
}

/** Splits one cell's already-laid-out paragraph boxes at `cutHeightPx`, never mid-line. */
function splitCellAtHeight(boxes: LayoutParagraphBox[], cutHeightPx: number): CellSplit {
	const before: LayoutParagraphBox[] = [];
	const after: LayoutParagraphBox[] = [];
	let beforeHeightPx = 0;
	for (const box of boxes) {
		if (box.yPx + box.heightPx <= cutHeightPx) {
			before.push(box);
			beforeHeightPx = Math.max(beforeHeightPx, box.yPx + box.heightPx);
			continue;
		}
		const beforeLines = box.lines.filter(
			(line) => box.yPx + line.yPx + line.heightPx <= cutHeightPx,
		);
		const afterLines = box.lines.slice(beforeLines.length);
		if (beforeLines.length) {
			const bottom = beforeLines.at(-1)!.yPx + beforeLines.at(-1)!.heightPx;
			before.push({ ...box, heightPx: bottom, lines: beforeLines });
			beforeHeightPx = Math.max(beforeHeightPx, box.yPx + bottom);
		}
		if (afterLines.length) {
			const shift = expectDefined(afterLines[0], 'first line after split').yPx;
			after.push({
				...box,
				yPx: box.yPx + shift,
				heightPx: afterLines.at(-1)!.yPx + afterLines.at(-1)!.heightPx - shift,
				lines: afterLines.map((line) => ({ ...line, yPx: line.yPx - shift })),
			});
		}
	}
	// Nested cells have paragraphs running in parallel, rather than a single sequential story.
	const afterOrigin = Math.min(...after.map((box) => box.yPx));
	for (const box of after) box.yPx -= afterOrigin;
	return { before, after, beforeHeightPx };
}

export interface RowSplit {
	before: RowLayout;
	after: RowLayout | null;
}

/**
 * Splits a row that does not fit in `cutHeightPx` of remaining page/column
 * space, including bottom padding. Cells normally share the same cut; optional
 * individual cuts defer a marker without truncating neighboring unreferenced cells.
 * Splits never cut a line, and continuation fragments retain their top and bottom padding.
 * Returns `after: null` when every cell fit (nothing left to continue).
 */
export function splitRowAtHeight(
	row: RowLayout,
	cutHeightPx: number,
	cellCuts?: readonly number[],
): RowSplit {
	const splits = row.cells.map((cell, index) => {
		const padding = row.cellPadding?.[index] ?? { top: 0, bottom: 0 };
		const split = splitCellAtHeight(cell, (cellCuts?.[index] ?? cutHeightPx) - padding.bottom);
		if (split.before.length) split.beforeHeightPx += padding.bottom;
		if (split.after.length)
			split.after = split.after.map((box) => ({ ...box, yPx: box.yPx + padding.top }));
		return split;
	});
	const anyAfter = splits.some((split) => split.after.length > 0);
	const before: RowLayout = {
		heightPx: Math.max(0, ...splits.map((split) => split.beforeHeightPx)),
		cells: splits.map((split) => split.before),
		geometry: row.geometry.map((geometry, index) => ({
			...geometry,
			contentHeightPx: splits[index]!.beforeHeightPx,
		})),
		...(row.cellPadding ? { cellPadding: row.cellPadding } : {}),
	};
	if (!anyAfter) return { before, after: null };
	const after: RowLayout = {
		heightPx: Math.max(
			0,
			...splits.map((split, index) =>
				split.after.length
					? split.after.at(-1)!.yPx +
						split.after.at(-1)!.heightPx +
						(row.cellPadding?.[index]?.bottom ?? 0)
					: 0,
			),
		),
		cells: splits.map((split) => split.after),
		geometry: row.geometry.map((geometry, index) => ({
			...geometry,
			contentHeightPx: splits[index]!.after.length
				? splits[index]!.after.at(-1)!.yPx +
					splits[index]!.after.at(-1)!.heightPx +
					(row.cellPadding?.[index]?.bottom ?? 0)
				: 0,
		})),
		...(row.cellPadding ? { cellPadding: row.cellPadding } : {}),
	};
	return { before, after };
}
