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
		const x =
			block.alignment === 'center'
				? Math.max(0, (width - tableWidth) / 2)
				: block.alignment === 'right'
					? Math.max(0, width - tableWidth)
					: (block.indentPx ?? 0);
		const frame: LayoutTableBox = {
			kind: 'table',
			blockId: block.id,
			xPx: x,
			widthPx: tableWidth,
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
			const layout = layoutRow(row, tableWidth, measurer, note, undefined, starts);
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
						yPx: y + paragraph.yPx,
						xPx: x + geometry.xPx + geometry.paddingLeftPx + (paragraph.xPx ?? 0),
						widthPx:
							paragraph.widthPx ??
							Math.max(1, geometry.widthPx - geometry.paddingLeftPx - geometry.paddingRightPx),
						readOnly: true,
					});
			});
			y += layout.heightPx;
		}
		frame.heightPx = y - frame.yPx;
		if (frame.rows.length) tables.push(frame);
	}
	// Cell paragraphs can run in parallel across a nested row; shared split machinery expects y order.
	boxes.sort((a, b) => a.yPx - b.yPx);
	return { heightPx: y, boxes, tables };
}
