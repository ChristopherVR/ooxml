// Floating tables (`w:tblpPr`, ECMA-376 17.4.58): placed out of the text flow at their frame
// position, like floating pictures, without taking room from the flow; the text of their page
// wraps around them on the next layout pass (wrap.ts).
import { floatPosition } from './floats';
import { layoutBorderedRow, TableFragmentBorders } from './table-border-layout';
import { nestedFrameProfile } from './nested-frame-profile';
import { cellParagraphs } from './table-cell-layout';
import type { PageCursor } from './page-cursor';
import type { TextMeasurer } from './measure';
import type { LayoutTable } from './input';
import type { LayoutTableBox, LayoutTableRowBox } from './result';

export const FLOATING_TABLE_NOTE =
	'Floating tables are placed whole on the page where they occur and text wraps around their rectangle; narrow gaps beside them stay empty.';
export const FLOATING_TABLE_IN_FLOW_NOTE =
	'Floating tables that do not fit on their page, hold footnote references or merge cells vertically are laid out in the text flow.';

/** Floating placement keeps to tables it lays out whole: no notes and no vertical merges. */
function floatable(table: LayoutTable): boolean {
	return table.rows.every((row) =>
		row.cells.every(
			(cell) =>
				!cell.verticalMerge &&
				!(cell.rowSpan && cell.rowSpan > 1) &&
				cellParagraphs(cell).every((paragraph) => !paragraph.footnotes?.length),
		),
	);
}

/**
 * Places a floating table on the current page at its frame position, without advancing the flow.
 * Returns false, placing nothing, for a table that must be laid out in the flow instead.
 */
export function placeFloatingTable(
	cursor: PageCursor,
	table: LayoutTable & { floating: NonNullable<LayoutTable['floating']> },
	measurer: TextMeasurer,
	note: (message: string) => void,
): boolean {
	if (!floatable(table)) {
		note(FLOATING_TABLE_IN_FLOW_NOTE);
		return false;
	}
	const widthPx = table.widthPx ?? cursor.columnWidthPx;
	// Rows reserve their borders as a body table's do: shared edges once, the last bottom edge kept.
	const profile = nestedFrameProfile(table, widthPx, true);
	const borders = new TableFragmentBorders();
	const rows: LayoutTableRowBox[] = [];
	let heightPx = 0;
	for (const row of table.rows) {
		const layout = layoutBorderedRow(row, widthPx, measurer, note, profile);
		heightPx += borders.remove(rows.at(-1));
		rows.push({
			yPx: heightPx,
			heightPx: layout.heightPx,
			repeated: false,
			cells: layout.cells,
			geometry: layout.geometry,
			...(layout.cellTables ? { cellTables: layout.cellTables } : {}),
		});
		heightPx += layout.heightPx;
		borders.record(layout.borderBottomPx);
	}
	heightPx += borders.restore(rows.at(-1));
	if (!profile) {
		// Without a measured border profile the rows hold no border room; the frame still includes
		// the outer top and bottom edges, which text must clear.
		const edge = (row: LayoutTable['rows'][number] | undefined, side: 'top' | 'bottom') =>
			Math.max(0, ...(row?.cells ?? []).map((cell) => cell.borders?.[side]?.widthPx ?? 0));
		heightPx += edge(table.rows[0], 'top') + edge(table.rows.at(-1), 'bottom');
	}
	const page = cursor.page;
	const column = cursor.column;
	const frame = table.floating;
	const position = floatPosition(
		{
			widthPx,
			heightPx,
			relativeFromH: frame.relativeFromH,
			...(frame.alignH ? { alignH: frame.alignH } : {}),
			offsetXPx: frame.offsetXPx,
			relativeFromV: frame.relativeFromV,
			...(frame.alignV ? { alignV: frame.alignV } : {}),
			offsetYPx: frame.offsetYPx,
		},
		page,
		column,
		{ topPx: page.marginTopPx + cursor.y, heightPx },
	);
	const topPx = position.yPx;
	if (topPx < 0 || topPx + heightPx > page.heightPx - page.marginBottomPx) {
		note(FLOATING_TABLE_IN_FLOW_NOTE);
		return false;
	}
	note(FLOATING_TABLE_NOTE);
	const box: LayoutTableBox = {
		kind: 'table',
		blockId: table.id,
		xPx: position.xPx - page.marginLeftPx - column.xPx,
		widthPx,
		yPx: 0,
		heightPx,
		rows,
		floating: {
			topPx: frame.distance.topPx,
			rightPx: frame.distance.rightPx,
			bottomPx: frame.distance.bottomPx,
			leftPx: frame.distance.leftPx,
		},
	};
	cursor.place(box, 0, topPx - page.marginTopPx - cursor.y);
	return true;
}
