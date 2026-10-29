import { reflowNestedFragments } from './nested-fragment-edges';
import { cellContentBottom } from './table-content';
import { clipTableFrames } from './table-frames';
import type { TextMeasurer } from './measure';
import { layoutParagraph } from './paragraph-layout';
import type { LayoutParagraph, LayoutTableRow } from './input';
import type { LayoutCellGeometry, LayoutParagraphBox, LayoutTableBox } from './result';
import { expectDefined } from '../expect-defined';
import { stackCellBlocks } from './table-cell-layout';

export interface RowLayout {
	heightPx: number;
	cells: LayoutParagraphBox[][];
	geometry: LayoutCellGeometry[];
	cellTables?: LayoutTableBox[][];
	/** Final physical edge reserved until a following row shares it. */
	borderBottomPx?: number;
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
	frameBorders = false,
	frameLeftBorderPx?: number,
	physicalFragments = reflowNestedFragments(row),
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

		const margins = cell.padding ?? { top: 0, right: 0, bottom: 0, left: 0 };
		const leftBorder = frameBorders ? (cell.borders?.left?.widthPx ?? 0) : 0;
		const rightBorder = frameBorders ? (cell.borders?.right?.widthPx ?? 0) : 0;
		const padding = frameBorders
			? {
					top: margins.top + (cell.borders?.top?.widthPx ?? 0),
					bottom: margins.bottom,
					left: leftBorder / 2 + Math.max(margins.left, leftBorder / 2),
					right: Math.max(margins.right, rightBorder / 2) - rightBorder / 2,
				}
			: margins;
		const externalRight =
			frameBorders && geometry.length === row.cells.length - 1 ? rightBorder : 0;
		const paintWidth =
			widthPx + (frameBorders ? (leftBorder - rightBorder) / 2 + externalRight : 0);
		const frameX = frameBorders ? ((frameLeftBorderPx ?? leftBorder) - leftBorder) / 2 : 0;
		const content = stackCellBlocks(
			cell,
			Math.max(1, paintWidth - padding.left - padding.right - externalRight),
			measurer,
			note,
			padding.top,
			starts,
			physicalFragments,
		);
		const heightPx = content.heightPx + padding.bottom;
		geometry.push({
			xPx: (cell.xPx ?? x) + frameX,
			widthPx: paintWidth,
			...(frameBorders ? { borderInsetsReserved: true } : {}),
			paddingLeftPx: padding.left,
			paddingRightPx: padding.right + externalRight,
			contentHeightPx: heightPx,
			...(cell.borders ? { borders: cell.borders } : {}),
			...(cell.shading ? { shading: cell.shading } : {}),
			...(cell.verticalAlign ? { verticalAlign: cell.verticalAlign } : {}),
			...(cell.rowSpan ? { mergeRows: cell.rowSpan } : {}),
		});
		x = (cell.xPx ?? x) + widthPx;
		return { heightPx, boxes: content.boxes, tables: content.tables ?? [] };
	});
	const content = Math.max(0, ...laidOut.map((cell) => cell.heightPx));
	const heightPx =
		row.heightPx === undefined
			? content
			: row.heightRule === 'exact'
				? row.heightPx
				: Math.max(content, row.heightPx);
	if (frameBorders) for (const cell of geometry) cell.alignmentHeightPx = heightPx;
	return {
		heightPx,
		cells: laidOut.map((cell) => cell.boxes),
		geometry,
		...(laidOut.some((cell) => cell.tables.length)
			? { cellTables: laidOut.map((cell) => cell.tables) }
			: {}),
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
	afterOrigin?: number;
}

/** Splits one cell's already-laid-out paragraph boxes at `cutHeightPx`, never mid-line. */
function splitCellAtHeight(
	boxes: LayoutParagraphBox[],
	cutHeightPx: number,
	paragraphCuts?: ReadonlyMap<string, number>,
): CellSplit {
	const before: LayoutParagraphBox[] = [];
	const after: LayoutParagraphBox[] = [];
	let beforeHeightPx = 0;
	for (const box of boxes) {
		const cut = paragraphCuts?.get(box.blockId) ?? cutHeightPx;
		if (
			box.yPx + box.heightPx + (box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0) <=
			cut
		) {
			before.push(box);
			beforeHeightPx = Math.max(
				beforeHeightPx,
				box.yPx + box.heightPx + (box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0),
			);
			continue;
		}
		const beforeLines = box.lines.filter(
			(line) =>
				box.yPx +
					line.yPx +
					line.heightPx +
					(box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0) <=
				cut,
		);
		const afterLines = box.lines.slice(beforeLines.length);
		if (beforeLines.length) {
			const bottom = beforeLines.at(-1)!.yPx + beforeLines.at(-1)!.heightPx;
			before.push({ ...box, heightPx: bottom, lines: beforeLines });
			beforeHeightPx = Math.max(
				beforeHeightPx,
				box.yPx + bottom + (box.tableFragmentBottomPx ?? box.tableFrameInsets?.bottom ?? 0),
			);
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
	const afterOrigin = Math.min(...after.map((box) => box.yPx - (box.tableFrameInsets?.top ?? 0)));
	for (const box of after) box.yPx -= afterOrigin;
	return { before, after, beforeHeightPx, ...(after.length ? { afterOrigin } : {}) };
}

export interface RowSplit {
	before: RowLayout;
	after: RowLayout | null;
}

/**
 * Splits a row that does not fit in `cutHeightPx` of remaining page/column
 * space, including bottom padding. Cells normally share the same cut; optional
 * individual cuts defer a marker without truncating neighboring unreferenced cells.
 * Paragraph-specific cuts are used for independently continued inner stories; their tail
 * must be source-reflowed by the table flow before placement to restore independent origins.
 * Splits never cut a line, and continuation fragments retain their top and bottom padding.
 * Returns `after: null` when every cell fit (nothing left to continue).
 */
export function splitRowAtHeight(
	row: RowLayout,
	cutHeightPx: number,
	cellCuts?: readonly number[],
	paragraphCuts?: ReadonlyMap<string, number>,
): RowSplit {
	const splits = row.cells.map((cell, index) => {
		const padding = row.cellPadding?.[index] ?? { top: 0, bottom: 0 };
		const cuts = paragraphCuts
			? new Map(
					cell.flatMap((box) => {
						const cut = paragraphCuts.get(box.blockId);
						return cut === undefined ? [] : [[box.blockId, cut - padding.bottom] as const];
					}),
				)
			: undefined;
		const split = splitCellAtHeight(
			cell,
			(cellCuts?.[index] ?? cutHeightPx) - padding.bottom,
			cuts,
		);
		if (split.before.length) split.beforeHeightPx += padding.bottom;
		if (split.after.length)
			split.after = split.after.map((box) => ({ ...box, yPx: box.yPx + padding.top }));
		return split;
	});
	const anyAfter = splits.some((split) => split.after.length > 0);
	const beforeHeightPx = Math.max(0, ...splits.map((split) => split.beforeHeightPx));
	const before: RowLayout = {
		heightPx: beforeHeightPx,
		cells: splits.map((split) => split.before),
		...(row.cellTables
			? {
					cellTables: row.cellTables.map((tables, index) =>
						clipTableFrames(
							tables,
							0,
							splits[index]!.beforeHeightPx - (row.cellPadding?.[index]?.bottom ?? 0),
						),
					),
				}
			: {}),
		geometry: row.geometry.map(({ mergeRows: _merge, ...geometry }, index) => ({
			...geometry,
			contentHeightPx: splits[index]!.beforeHeightPx,
			...(row.borderBottomPx !== undefined ? { alignmentHeightPx: beforeHeightPx } : {}),
		})),
		...(row.cellPadding ? { cellPadding: row.cellPadding } : {}),
		...(row.borderBottomPx !== undefined ? { borderBottomPx: row.borderBottomPx } : {}),
	};
	if (!anyAfter) return { before, after: null };
	const afterHeightPx = Math.max(
		0,
		...splits.map((split, index) =>
			cellContentBottom(split.after, row.cellPadding?.[index]?.bottom ?? 0),
		),
	);
	const after: RowLayout = {
		heightPx: afterHeightPx,
		cells: splits.map((split) => split.after),
		...(row.cellTables
			? {
					cellTables: row.cellTables.map((tables, index) =>
						splits[index]!.after.length
							? clipTableFrames(
									tables,
									splits[index]!.afterOrigin ?? 0,
									row.heightPx,
									row.cellPadding?.[index]?.top ?? 0,
								)
							: [],
					),
				}
			: {}),
		geometry: row.geometry.map(({ mergeRows: _merge, ...geometry }, index) => ({
			...geometry,
			contentHeightPx: cellContentBottom(
				splits[index]!.after,
				row.cellPadding?.[index]?.bottom ?? 0,
			),
			...(row.borderBottomPx !== undefined ? { alignmentHeightPx: afterHeightPx } : {}),
		})),
		...(row.cellPadding ? { cellPadding: row.cellPadding } : {}),
		...(row.borderBottomPx !== undefined ? { borderBottomPx: row.borderBottomPx } : {}),
	};
	return { before, after };
}
