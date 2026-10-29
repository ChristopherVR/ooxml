import { reflowNestedFragments } from './nested-fragment-edges';
import { cellContentBottom } from './table-content';
import type { TextMeasurer } from './measure';
import { layoutParagraph } from './paragraph-layout';
import type { LayoutParagraph, LayoutTableRow } from './input';
import type { LayoutCellGeometry, LayoutParagraphBox, LayoutTableBox } from './result';
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
	/** Space after the last paragraph of each cell (px): it belongs to the end of the story, so a
	 * continuation fragment keeps it even though the paragraph boxes do not carry it. */
	trailingPx?: number[];
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
		const trailingPx = content.tables?.length
			? 0
			: Math.max(0, content.heightPx - cellContentBottom(content.boxes));
		return { heightPx, boxes: content.boxes, tables: content.tables ?? [], trailingPx };
	});
	// The text of a vertically merged cell is sized with the rows it spans, not with its first row.
	const content = Math.max(
		0,
		...laidOut.map((cell, index) =>
			(row.cells[index]?.rowSpan ?? 1) > 1 && !row.mergedTextInRow ? 0 : cell.heightPx,
		),
	);
	const heightPx =
		row.heightPx === undefined
			? content
			: row.heightRule === 'exact'
				? row.heightPx
				: Math.max(content, row.heightPx);
	if (frameBorders) for (const cell of geometry) cell.alignmentHeightPx = heightPx;
	const trailing = laidOut.map((cell) => cell.trailingPx);
	return {
		heightPx,
		cells: laidOut.map((cell) => cell.boxes),
		geometry,
		...(trailing.some((px) => px > 0) ? { trailingPx: trailing } : {}),
		...(laidOut.some((cell) => cell.tables.length)
			? { cellTables: laidOut.map((cell) => cell.tables) }
			: {}),
		cellPadding: row.cells.map((cell) => ({
			top: cell.padding?.top ?? 0,
			bottom: cell.padding?.bottom ?? 0,
		})),
	};
}

export { splitRowAtHeight, type RowSplit } from './table-row-split';
