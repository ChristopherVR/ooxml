import { nestedFrameBorderWidth } from './nested-frame-profile';
import { cellVerticalOffset } from './result';
import type { LayoutBlock, LayoutParagraph, LayoutTableCell } from './input';
import type { LayoutParagraphBox, LayoutTableBox } from './result';
import type { TextMeasurer } from './measure';
import { layoutRow, stackParagraphs } from './table-layout';
/** All nested source paragraphs, in OOXML story order. */
export function cellParagraphs(cell: LayoutTableCell): LayoutParagraph[] {
	const walk = (blocks: LayoutBlock[]): LayoutParagraph[] =>
		blocks.flatMap((block) =>
			block.kind === 'paragraph'
				? [block]
				: block.rows.flatMap((row) => row.cells.flatMap(cellParagraphs)),
		);
	return cell.blocks ? walk(cell.blocks) : cell.paragraphs;
}
/** Project read-only nested cell stories using the same paragraph and row layout as body tables. */
export function stackCellBlocks(
	cell: LayoutTableCell,
	width: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
	top: number,
	starts?: ReadonlyMap<string, number>,
): { heightPx: number; boxes: LayoutParagraphBox[]; tables?: LayoutTableBox[] } {
	if (!cell.blocks)
		return stackParagraphs(
			starts ? cell.paragraphs.filter((p) => starts.has(p.id)) : cell.paragraphs,
			width,
			measurer,
			note,
			top,
			starts,
		);
	let y = top;
	const boxes: LayoutParagraphBox[] = [];
	const tables: LayoutTableBox[] = [];
	for (const block of cell.blocks) {
		if (block.kind === 'paragraph') {
			if (starts && !starts.has(block.id)) continue;
			const result = stackParagraphs([block], width, measurer, note, y, starts);
			boxes.push(...result.boxes);
			y = result.heightPx;
			continue;
		}
		const tableWidth = block.widthPx ?? width;
		const profileWidth = nestedFrameBorderWidth(block);
		const reserveBorders = profileWidth !== undefined;
		if (!reserveBorders)
			note(
				'Nested table border occupancy is approximate for mixed widths, missing edges or non-solid borders; native evidence covers uniform single-line borders.',
			);
		const externalRight = profileWidth ?? 0;
		const paintedWidth = tableWidth + externalRight;
		const x =
			block.alignment === 'center'
				? Math.max(0, (width - paintedWidth) / 2)
				: block.alignment === 'right'
					? Math.max(0, width - paintedWidth)
					: (block.indentPx ?? 0);
		const frame: LayoutTableBox = {
			kind: 'table',
			blockId: block.id,
			xPx: x,
			widthPx: paintedWidth,
			yPx: y,
			heightPx: 0,
			rows: [],
		};
		for (const row of block.rows) {
			if (
				starts &&
				!row.cells.some((nested) => cellParagraphs(nested).some((p) => starts.has(p.id)))
			)
				continue;
			const layout = layoutRow(row, tableWidth, measurer, note, undefined, starts, reserveBorders);
			frame.rows.push({
				yPx: y - frame.yPx,
				heightPx: layout.heightPx,
				repeated: false,
				cells: layout.cells,
				geometry: layout.geometry,
				...(layout.cellTables ? { cellTables: layout.cellTables } : {}),
			});
			layout.cells.forEach((paragraphs, index) => {
				const geometry = layout.geometry[index]!;
				for (const paragraph of paragraphs)
					boxes.push({
						...paragraph,
						yPx: y + paragraph.yPx + cellVerticalOffset(geometry, layout.heightPx),
						xPx: x + geometry.xPx + geometry.paddingLeftPx + (paragraph.xPx ?? 0),
						widthPx:
							paragraph.widthPx ??
							Math.max(1, geometry.widthPx - geometry.paddingLeftPx - geometry.paddingRightPx),
						readOnly: true,
						tableFrameInsets: {
							top:
								(paragraph.tableFrameInsets?.top ?? 0) +
								(reserveBorders ? (geometry.borders?.top?.widthPx ?? 0) : 0),
							bottom:
								(paragraph.tableFrameInsets?.bottom ?? 0) +
								(reserveBorders ? (geometry.borders?.bottom?.widthPx ?? 0) : 0),
						},
					});
			});
			y += layout.heightPx;
		}
		const bottomBorder = profileWidth ?? 0;
		if (frame.rows.length) frame.rows.at(-1)!.heightPx += bottomBorder;
		y += bottomBorder;
		frame.heightPx = y - frame.yPx;
		if (frame.rows.length) tables.push(frame);
	}
	// Cell paragraphs can run in parallel across a nested row; shared split machinery expects y order.
	boxes.sort((a, b) => a.yPx - b.yPx);
	return { heightPx: y, boxes, tables };
}
