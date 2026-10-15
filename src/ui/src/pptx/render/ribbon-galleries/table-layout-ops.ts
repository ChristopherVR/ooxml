/**
 * Table Layout > Rows & Columns, Merge, Cell Size and Alignment as pure table edits: each takes
 * the table element and the selected cells and returns the element patch the binding applies
 * through its normal update + history path, or `null` when the edit does not apply.
 *
 * Structure edits go through the element-level helpers so `rawXml` stays in step with
 * `tableData` (the renderers and the save writer read it in preference). Size, merge and cell
 * style edits patch `tableData` only, as the Inspector's table panel does.
 *
 * @module render/ribbon-galleries/table-layout-ops
 */
import type { PptxTableCellStyle, PptxTableData, TablePptxElement } from 'ooxml-core/pptx';
import { updateMergeAttrsInRawXml } from 'ooxml-core/pptx';

import {
	insertTableElementColumn,
	insertTableElementRow,
	removeTableElementColumn,
	removeTableElementRow,
} from '../table-data-grid-ops';
import { canMergeCells, canSplitCell, mergeCells, splitCell } from '../table-merge';
import { evenColumnWidths, evenRowHeights } from '../table-resize';
import type { RibbonGalleryApplyResult } from './gallery-types';

/** A selected cell, as every binding tracks it. */
export interface TableLayoutCell {
	row: number;
	column: number;
}

export interface TableSelectionBounds {
	minRow: number;
	maxRow: number;
	minColumn: number;
	maxColumn: number;
}

/** The cursor / range state React, Vue and Angular keep for the table under edit. */
export interface TableCursorState {
	rowIndex: number;
	columnIndex: number;
	selectedCells?: ReadonlyArray<{ row: number; col: number }> | undefined;
}

/**
 * The cells a Table Layout command acts on: the selected range, else the cell under the cursor,
 * else none (the table frame is selected).
 */
export function tableLayoutCells(state: TableCursorState | null | undefined): TableLayoutCell[] {
	if (!state) {
		return [];
	}
	if (state.selectedCells && state.selectedCells.length > 0) {
		return state.selectedCells.map(({ row, col }) => ({ row, column: col }));
	}
	return [{ row: state.rowIndex, column: state.columnIndex }];
}

export function selectionBounds(cells: readonly TableLayoutCell[]): TableSelectionBounds | null {
	if (cells.length === 0) {
		return null;
	}
	return {
		minRow: Math.min(...cells.map((c) => c.row)),
		maxRow: Math.max(...cells.map((c) => c.row)),
		minColumn: Math.min(...cells.map((c) => c.column)),
		maxColumn: Math.max(...cells.map((c) => c.column)),
	};
}

function elementResult(
	before: TablePptxElement,
	after: TablePptxElement,
): RibbonGalleryApplyResult | null {
	if (after === before || !after.tableData) {
		return null;
	}
	const patch: Partial<TablePptxElement> = { tableData: after.tableData };
	if (after.rawXml !== before.rawXml) {
		patch.rawXml = after.rawXml as NonNullable<TablePptxElement['rawXml']>;
	}
	return { kind: 'element', elementId: before.id, patch };
}

function dataResult(
	element: TablePptxElement,
	next: PptxTableData | undefined,
	syncMerges = false,
): RibbonGalleryApplyResult | null {
	if (!next || next === element.tableData) {
		return null;
	}
	const patch: Partial<TablePptxElement> = { tableData: next };
	// Merge spans live in the graphic frame markup too, which the renderers read first.
	const rawXml = syncMerges ? updateMergeAttrsInRawXml(element, next) : undefined;
	if (rawXml) {
		patch.rawXml = rawXml as NonNullable<TablePptxElement['rawXml']>;
	}
	return { kind: 'element', elementId: element.id, patch };
}

const asCoords = (cells: readonly TableLayoutCell[]) =>
	cells.map((c) => ({ row: c.row, col: c.column }));

export type InsertSide = 'above' | 'below' | 'left' | 'right';

/**
 * Insert a row or column next to the selection. With no selection the table's edge is used
 * (above/left: before the first, below/right: after the last).
 */
export function insertNextToSelection(
	element: TablePptxElement,
	cells: readonly TableLayoutCell[],
	side: InsertSide,
): RibbonGalleryApplyResult | null {
	const data = element.tableData;
	if (!data) {
		return null;
	}
	const bounds = selectionBounds(cells);
	const lastRow = data.rows.length - 1;
	const lastColumn = data.columnWidths.length - 1;
	switch (side) {
		case 'above':
			return elementResult(element, insertTableElementRow(element, bounds?.minRow ?? 0, 'above'));
		case 'below':
			return elementResult(
				element,
				insertTableElementRow(element, bounds?.maxRow ?? lastRow, 'below'),
			);
		case 'left':
			return elementResult(
				element,
				insertTableElementColumn(element, bounds?.minColumn ?? 0, 'left'),
			);
		case 'right':
			return elementResult(
				element,
				insertTableElementColumn(element, bounds?.maxColumn ?? lastColumn, 'right'),
			);
	}
}

/** Delete every row (or column) the selection touches; the last one is never removed. */
export function deleteSelected(
	element: TablePptxElement,
	cells: readonly TableLayoutCell[],
	axis: 'row' | 'column',
): RibbonGalleryApplyResult | null {
	const bounds = selectionBounds(cells);
	if (!bounds) {
		return null;
	}
	let next = element;
	const from = axis === 'row' ? bounds.minRow : bounds.minColumn;
	const to = axis === 'row' ? bounds.maxRow : bounds.maxColumn;
	for (let index = to; index >= from; index -= 1) {
		next =
			axis === 'row' ? removeTableElementRow(next, index) : removeTableElementColumn(next, index);
	}
	return elementResult(element, next);
}

/** True when the selection is a mergeable rectangle of two or more cells. */
export function canMergeSelection(data: PptxTableData, cells: readonly TableLayoutCell[]): boolean {
	return cells.length > 1 && canMergeCells(asCoords(cells), data);
}

/** True when the selection holds the anchor of a merged cell. */
export function canSplitSelection(data: PptxTableData, cells: readonly TableLayoutCell[]): boolean {
	return cells.some((c) => canSplitCell(c.row, c.column, data));
}

export function mergeSelection(
	element: TablePptxElement,
	cells: readonly TableLayoutCell[],
): RibbonGalleryApplyResult | null {
	const data = element.tableData;
	if (!data || !canMergeSelection(data, cells)) {
		return null;
	}
	return dataResult(element, mergeCells(asCoords(cells), data), true);
}

export function splitSelection(
	element: TablePptxElement,
	cells: readonly TableLayoutCell[],
): RibbonGalleryApplyResult | null {
	let data = element.tableData;
	if (!data) {
		return null;
	}
	for (const cell of cells) {
		if (canSplitCell(cell.row, cell.column, data)) {
			data = splitCell(cell.row, cell.column, data);
		}
	}
	return dataResult(element, data, true);
}

/**
 * Distribute rows or columns evenly. Two or more selected rows (columns) are evened among
 * themselves; otherwise the whole table is, as PowerPoint does.
 */
export function distribute(
	element: TablePptxElement,
	cells: readonly TableLayoutCell[],
	axis: 'row' | 'column',
): RibbonGalleryApplyResult | null {
	const data = element.tableData;
	if (!data) {
		return null;
	}
	const bounds = selectionBounds(cells);
	if (axis === 'row') {
		const spans = bounds !== null && bounds.maxRow > bounds.minRow;
		const from = spans ? bounds.minRow : 0;
		const to = spans ? bounds.maxRow : data.rows.length - 1;
		const evened = evenRowHeights(data.rows.slice(from, to + 1));
		const rows = data.rows.map((row, index) =>
			index >= from && index <= to ? (evened[index - from] ?? row) : row,
		);
		return dataResult(element, { ...data, rows });
	}
	const count = data.columnWidths.length;
	const spans = bounds !== null && bounds.maxColumn > bounds.minColumn;
	const from = spans ? bounds.minColumn : 0;
	const to = spans ? bounds.maxColumn : count - 1;
	if (from === 0 && to === count - 1) {
		return dataResult(element, { ...data, columnWidths: evenColumnWidths(count) });
	}
	const slice = data.columnWidths.slice(from, to + 1);
	const share = slice.reduce((sum, width) => sum + width, 0) / slice.length;
	return dataResult(element, {
		...data,
		columnWidths: data.columnWidths.map((width, index) =>
			index >= from && index <= to ? share : width,
		),
	});
}

/**
 * Apply `patch` (and drop the `clear` keys) on the selected cells, or on every cell when the
 * table frame itself is selected (no cells).
 */
export function styleCells(
	element: TablePptxElement,
	cells: readonly TableLayoutCell[],
	patch: Partial<PptxTableCellStyle>,
	clear: ReadonlyArray<keyof PptxTableCellStyle> = [],
): RibbonGalleryApplyResult | null {
	const data = element.tableData;
	if (!data) {
		return null;
	}
	const keys = new Set(cells.map((c) => `${c.row}:${c.column}`));
	const all = keys.size === 0;
	const restyle = (style: PptxTableCellStyle | undefined): PptxTableCellStyle => {
		const next: PptxTableCellStyle = { ...style, ...patch };
		for (const key of clear) {
			delete next[key];
		}
		return next;
	};
	return dataResult(element, {
		...data,
		rows: data.rows.map((row, r) => ({
			...row,
			cells: row.cells.map((cell, c) =>
				all || keys.has(`${r}:${c}`) ? { ...cell, style: restyle(cell.style) } : cell,
			),
		})),
	});
}
