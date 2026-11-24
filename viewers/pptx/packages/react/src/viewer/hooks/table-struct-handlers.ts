/**
 * Table structure handlers: cell editing, column / row resize,
 * insert / delete rows and columns.
 *
 * Structural operations (insert/delete row/column) handle merge span
 * adjustments and synchronise both `tableData` and `rawXml` so that
 * rendering and saving both reflect the changes.
 */
import {
	insertTableElementRow,
	removeTableElementRow,
	insertTableElementColumn,
	removeTableElementColumn,
	getCellText,
	setCellText,
	resolveTableCellEditEnd,
	tableCellsChangeCasePatch,
	tableCellsTextStylePatch,
} from 'ooxml-ui/pptx';
import type { ChangeCaseMode, TableCellEditEnd, TableCellsFontUpdate } from 'ooxml-ui/pptx';

import type { TableCellEditorState } from '../types';
import { updateCellTextInRawXml } from '../utils/table-parse';
import type { UseTableOperationsInput, TableStructHandlers } from './table-operation-types';

// ---------------------------------------------------------------------------
// Handler factory
//
// The pure structural math (merge-aware row/column insert/delete over
// `PptxTableData`) now lives in `pptx-viewer-shared` (`render/table-layout.ts`).
// These handlers wrap those pure transforms with the editor's stateful concerns
// (rawXml synchronisation, history dirty-marking, element updates).
// ---------------------------------------------------------------------------

export function createTableStructHandlers(input: UseTableOperationsInput): TableStructHandlers {
	const {
		selectedElement,
		tableEditorState: ts,
		elementLookup,
		setTableEditorState,
		ops,
		history,
		transformCommittedText,
	} = input;

	// ── Cell text editing ─────────────────────────────────────────────────

	const handleCommitCellEdit = (
		elementId: string,
		rowIndex: number,
		colIndex: number,
		rawText: string,
		end?: TableCellEditEnd,
	) => {
		const el = elementLookup.get(elementId);
		if (!el || el.type !== 'table') {
			return;
		}
		if (end) {
			// The rich cell editor: shared `resolveTableCellEditEnd` decides the
			// document update (runs, appended row, grid paste) and the next cell.
			const outcome = resolveTableCellEditEnd(el, rowIndex, colIndex, end, {
				...(transformCommittedText ? { transformText: transformCommittedText } : {}),
			});
			if (outcome.element) {
				ops.updateElementById(elementId, {
					tableData: outcome.element.tableData,
					...(outcome.element.rawXml ? { rawXml: outcome.element.rawXml } : {}),
				});
				history.markDirty();
			}
			const target = outcome.edit ?? outcome.select ?? { row: rowIndex, col: colIndex };
			setTableEditorState({
				rowIndex: target.row,
				columnIndex: target.col,
				elementId,
				...(outcome.edit ? { isEditing: true, caret: outcome.edit.caret } : {}),
			} as TableCellEditorState);
			return;
		}

		const text = transformCommittedText
			? transformCommittedText(rawText, getCellText(el, rowIndex, colIndex))
			: rawText;
		const updates: Record<string, unknown> = {};

		// Always update tableData if it exists. Through shared `setCellText`, not
		// a local `{ ...cell, text }`: the cell also carries `textRuns`, the
		// per-run model the renderer prefers over the flat string, and a
		// hand-rolled spread keeps the runs describing the OLD wording - which is
		// exactly what the canvas then keeps painting.
		if (el.tableData) {
			updates.tableData = setCellText(el, rowIndex, colIndex, text).tableData;
		}

		// Always update rawXml if it exists (rendering reads from rawXml
		// via parseTableElementData, so it must stay in sync)
		if (el.rawXml) {
			const newRawXml = updateCellTextInRawXml(el, rowIndex, colIndex, text);
			if (newRawXml) {
				updates.rawXml = newRawXml;
			}
		}

		if (Object.keys(updates).length === 0) {
			return;
		}
		ops.updateElementById(elementId, updates);
		history.markDirty();
		setTableEditorState({
			rowIndex,
			columnIndex: colIndex,
			elementId,
		} as TableCellEditorState);
	};

	// ── Cell text style update ───────────────────────────────────────────

	// Every run of every selected cell (the Shift+click range, else the active
	// cell; the whole table when no cell is active), through the shared core
	// op so `tableData`, the painted runs and `rawXml` change together.
	const selectedTableCells = () =>
		ts
			? ts.selectedCells?.length
				? ts.selectedCells
				: [{ row: ts.rowIndex, col: ts.columnIndex }]
			: undefined;

	const handleUpdateCellTextStyle = (styleUpdates: Record<string, unknown>) => {
		if (!selectedElement || selectedElement.type !== 'table') {
			return;
		}
		const patch = tableCellsTextStylePatch(
			selectedElement,
			selectedTableCells(),
			styleUpdates as TableCellsFontUpdate,
		);
		if (Object.keys(patch).length === 0) {
			return;
		}
		ops.updateElementById(selectedElement.id, patch);
		history.markDirty();
	};

	const handleTransformCellTextCase = (mode: ChangeCaseMode) => {
		if (!selectedElement || selectedElement.type !== 'table') {
			return;
		}
		const patch = tableCellsChangeCasePatch(selectedElement, selectedTableCells(), mode);
		if (Object.keys(patch).length === 0) {
			return;
		}
		ops.updateElementById(selectedElement.id, patch);
		history.markDirty();
	};

	// ── Column / row resizing ─────────────────────────────────────────────

	const handleResizeTableColumns = (elementId: string, newWidths: number[]) => {
		const el = elementLookup.get(elementId);
		if (!el || el.type !== 'table' || !el.tableData) {
			return;
		}
		ops.updateElementById(elementId, {
			tableData: { ...el.tableData, columnWidths: newWidths },
		});
		history.markDirty();
	};

	const handleResizeTableRow = (elementId: string, rowIndex: number, newHeight: number) => {
		const el = elementLookup.get(elementId);
		if (!el || el.type !== 'table' || !el.tableData) {
			return;
		}
		const newRows = el.tableData.rows.map((row, i) =>
			i === rowIndex ? { ...row, height: newHeight } : row,
		);
		ops.updateElementById(elementId, {
			tableData: { ...el.tableData, rows: newRows },
		});
		history.markDirty();
	};

	// ── Insert row ────────────────────────────────────────────────────────

	const handleInsertTableRow = (position: 'above' | 'below') => {
		if (!selectedElement || selectedElement.type !== 'table' || !selectedElement.tableData) {
			return;
		}
		const rowIdx = ts?.rowIndex ?? 0;
		const next = insertTableElementRow(selectedElement, rowIdx, position);
		ops.updateSelectedElement({ tableData: next.tableData, rawXml: next.rawXml });
		history.markDirty();
	};

	// ── Delete row ────────────────────────────────────────────────────────

	const handleDeleteTableRow = () => {
		if (!selectedElement || selectedElement.type !== 'table' || !selectedElement.tableData) {
			return;
		}
		const rowIdx = ts?.rowIndex ?? 0;
		const next = removeTableElementRow(selectedElement, rowIdx);
		if (next === selectedElement) {
			return;
		}
		ops.updateSelectedElement({ tableData: next.tableData, rawXml: next.rawXml });
		history.markDirty();
	};

	// ── Insert column ─────────────────────────────────────────────────────

	const handleInsertTableColumn = (position: 'left' | 'right') => {
		if (!selectedElement || selectedElement.type !== 'table' || !selectedElement.tableData) {
			return;
		}
		const colIdx = ts?.columnIndex ?? 0;
		const next = insertTableElementColumn(selectedElement, colIdx, position);
		ops.updateSelectedElement({ tableData: next.tableData, rawXml: next.rawXml });
		history.markDirty();
	};

	// ── Delete column ─────────────────────────────────────────────────────

	const handleDeleteTableColumn = () => {
		if (!selectedElement || selectedElement.type !== 'table' || !selectedElement.tableData) {
			return;
		}
		const colIdx = ts?.columnIndex ?? 0;
		const next = removeTableElementColumn(selectedElement, colIdx);
		if (next === selectedElement) {
			return;
		}
		ops.updateSelectedElement({ tableData: next.tableData, rawXml: next.rawXml });
		history.markDirty();
	};

	return {
		handleCommitCellEdit,
		handleUpdateCellTextStyle,
		handleTransformCellTextCase,
		handleResizeTableColumns,
		handleResizeTableRow,
		handleInsertTableRow,
		handleDeleteTableRow,
		handleInsertTableColumn,
		handleDeleteTableColumn,
	};
}
