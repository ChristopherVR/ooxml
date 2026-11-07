import { rebuildTableStructureInRawXml } from '../../index';
import type { PptxTableData, TablePptxElement, TableStructureEdit } from '../../index';
import {
	deleteTableColumn,
	deleteTableRow,
	insertTableColumn,
	insertTableRow,
} from '../../editor/render/table-layout';

import type { ToolContext, ToolResult } from '../types';
import { validateSlideIndex } from './helpers';

// ── updateTableCells ─────────────────────────────────────────────────────────

export interface UpdateTableCellsParams {
	slideIndex: number;
	elementId: string;
	cells: Array<{ row: number; col: number; text: string }>;
}

export function updateTableCells(
	ctx: ToolContext,
	params: UpdateTableCellsParams,
): ToolResult<{ updatedCount: number }> {
	const err = validateSlideIndex(params.slideIndex, ctx.pptxData.slides.length);
	if (err) {
		throw new Error(err);
	}

	const slide = ctx.pptxData.slides[params.slideIndex];
	const el = slide.elements.find((e) => e.id === params.elementId);
	if (!el) {
		throw new Error(`Element '${params.elementId}' not found on slide ${params.slideIndex}.`);
	}
	if (el.type !== 'table') {
		throw new Error(`Element '${params.elementId}' is not a table.`);
	}

	const tbl = el as TablePptxElement;
	if (!tbl.tableData) {
		throw new Error(`Table '${params.elementId}' has no tableData.`);
	}

	let count = 0;
	for (const cellUpdate of params.cells) {
		const row = tbl.tableData.rows[cellUpdate.row];
		if (!row) {
			throw new Error(
				`Row ${cellUpdate.row} out of range (table has ${tbl.tableData.rows.length} rows).`,
			);
		}
		const cell = row.cells[cellUpdate.col];
		if (!cell) {
			throw new Error(`Column ${cellUpdate.col} out of range (row has ${row.cells.length} cells).`);
		}
		cell.text = cellUpdate.text;
		count++;
	}

	return {
		pptxData: ctx.pptxData,
		dirty: true,
		result: { updatedCount: count },
	};
}

// ── manageTableStructure ──────────────────────────────────────────────────────

export interface ManageTableStructureParams {
	slideIndex: number;
	elementId: string;
	action: 'insertRow' | 'deleteRow' | 'insertColumn' | 'deleteColumn';
	position?: number;
	referenceIndex?: number;
	cellTexts?: string[];
}

export function manageTableStructure(
	ctx: ToolContext,
	params: ManageTableStructureParams,
): ToolResult<{ rowCount: number; columnCount: number }> {
	const err = validateSlideIndex(params.slideIndex, ctx.pptxData.slides.length);
	if (err) {
		throw new Error(err);
	}

	const slide = ctx.pptxData.slides[params.slideIndex];
	const el = slide.elements.find((e) => e.id === params.elementId);
	if (!el) {
		throw new Error(`Element '${params.elementId}' not found on slide ${params.slideIndex}.`);
	}
	if (el.type !== 'table') {
		throw new Error(`Element '${params.elementId}' is not a table.`);
	}

	const tbl = el as TablePptxElement;
	if (!tbl.tableData) {
		throw new Error(`Table '${params.elementId}' has no tableData.`);
	}

	const tableData = tbl.tableData;
	const { rows, columnWidths } = tableData;
	const colCount = columnWidths.length;
	const rowCount = rows.length;
	let edit: TableStructureEdit;
	let next: PptxTableData;

	switch (params.action) {
		case 'insertRow': {
			const index = Math.min(Math.max(params.position ?? rowCount, 0), rowCount);
			edit = { axis: 'row', action: 'insert', index };
			next = insertTableRow(tableData, index, 'above', params.cellTexts);
			break;
		}
		case 'deleteRow': {
			if (rowCount <= 1) {
				throw new Error('Cannot delete the last row of a table.');
			}
			const index = params.referenceIndex ?? rowCount - 1;
			if (index < 0 || index >= rowCount) {
				throw new Error(`Row index ${index} out of range (0–${rowCount - 1}).`);
			}
			edit = { axis: 'row', action: 'delete', index };
			next = deleteTableRow(tableData, index);
			break;
		}
		case 'insertColumn': {
			const index = Math.min(Math.max(params.position ?? colCount, 0), colCount);
			edit = { axis: 'column', action: 'insert', index };
			next = insertTableColumn(tableData, index, 'left', params.cellTexts);
			// Keep the tool's width policy: reserve an equal share for the new column.
			next.columnWidths = columnWidths.map((width) => width * (colCount / (colCount + 1)));
			next.columnWidths.splice(index, 0, 1 / (colCount + 1));
			break;
		}
		case 'deleteColumn': {
			if (colCount <= 1) {
				throw new Error('Cannot delete the last column of a table.');
			}
			const index = params.referenceIndex ?? colCount - 1;
			if (index < 0 || index >= colCount) {
				throw new Error(`Column index ${index} out of range (0–${colCount - 1}).`);
			}
			edit = { axis: 'column', action: 'delete', index };
			next = deleteTableColumn(tableData, index);
			const share = columnWidths[index] / (colCount - 1);
			next.columnWidths = columnWidths
				.filter((_, column) => column !== index)
				.map((width) => width + share);
			break;
		}
		default: {
			throw new Error(`Unknown action: ${String(params.action)}`);
		}
	}

	const rawXml = rebuildTableStructureInRawXml(tbl, next, edit);
	tbl.tableData = next;
	if (rawXml) {
		tbl.rawXml = rawXml;
	}
	return {
		pptxData: ctx.pptxData,
		dirty: true,
		result: { rowCount: next.rows.length, columnCount: next.columnWidths.length },
	};
}
