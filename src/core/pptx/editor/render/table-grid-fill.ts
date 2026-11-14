/**
 * table-grid-fill.ts: write a block of plain-text values into a table,
 * growing it first when the block runs past the last row or column. This is
 * the document half of pasting tab-separated text (cells copied from a
 * spreadsheet as plain text) into a table; the parse and the plan of which
 * cells to write live in the UI layer.
 *
 * Each written cell takes the text in its own formatting (the destination
 * formatting PowerPoint uses for plain text): `setTableElementCellText`
 * keeps the cell's first run properties in `rawXml` and its cell style.
 *
 * @module render/table-grid-fill
 */
import type { TablePptxElement } from 'ooxml-core/pptx';

import {
	appendTableElementRow,
	insertTableElementColumn,
	setTableElementCellText,
} from './table-data-grid-ops';

/** One value to write, in grid coordinates of the grown table. */
export interface TableGridWrite {
	row: number;
	col: number;
	text: string;
}

/** Rows and columns to append, then the cells to write. */
export interface TableGridFill {
	addRows: number;
	addColumns: number;
	writes: readonly TableGridWrite[];
}

/**
 * Apply a grid fill as one element update: append `addRows` rows below the
 * last row and `addColumns` columns right of the last column (each copying
 * its neighbour's formatting, as Insert Below / Insert Right do), then write
 * every value. Returns the element itself when nothing changes.
 */
export function fillTableCells(element: TablePptxElement, fill: TableGridFill): TablePptxElement {
	let next = element;
	for (let index = 0; index < fill.addRows; index++) {
		next = appendTableElementRow(next);
	}
	for (let index = 0; index < fill.addColumns; index++) {
		const columns = next.tableData?.columnWidths.length ?? 0;
		if (columns === 0) {
			break;
		}
		next = insertTableElementColumn(next, columns - 1, 'right');
	}
	for (const write of fill.writes) {
		const cell = next.tableData?.rows[write.row]?.cells[write.col];
		if (!cell || cell.hMerge || cell.vMerge || cell.text === write.text) {
			continue;
		}
		next = setTableElementCellText(next, write.row, write.col, write.text);
	}
	return next;
}
