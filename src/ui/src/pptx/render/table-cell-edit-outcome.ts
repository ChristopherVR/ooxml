/**
 * table-cell-edit-outcome.ts: what happens to the document and the selection
 * when a table cell editor closes, decided once for every binding.
 *
 * A binding hands the editor's `TableCellEditEnd` here and maps the result:
 * replace the element with `element` (one history entry) when it is set,
 * then edit `edit` (Tab, Shift+Tab, an arrow at the text's edge), select
 * `select` without editing (Escape, a grid paste), or simply stop editing
 * (blur).
 *
 * @module render/table-cell-edit-outcome
 */
import type { PptxTableCell, TablePptxElement, TextSegment } from 'ooxml-core/pptx';
import { appendTableElementRow } from 'ooxml-core/pptx/editor/render/table-data-grid-ops';
import {
	editSegmentsText,
	segmentsPlainText,
} from 'ooxml-core/pptx/editor/render/segments-text-edit';
import { fillTableCells } from 'ooxml-core/pptx/editor/render/table-grid-fill';
import {
	setTableElementCellSegments,
	tableCellSegments,
} from 'ooxml-core/pptx/editor/render/table-cell-rich-text';

import type { TableCellCaret } from './table-cell-navigation';
import { resolveTableCellNavigation } from './table-cell-navigation';
import type { TableCellEditEnd } from './table-cell-text-editor';
import { isTabularPaste, parseTabularText, planTableGridPaste } from './table-tsv-paste';

export interface TableCellEditOutcome {
	/** The updated element, when the document changed. */
	element?: TablePptxElement;
	/** The cell to edit next, with where its caret starts. */
	edit?: { row: number; col: number; caret: TableCellCaret };
	/** The cell to leave selected (not editing). */
	select?: { row: number; col: number };
}

export interface TableCellEditOptions {
	/**
	 * The host's committed-text transform (Options > Proofing > AutoCorrect),
	 * given the edited text and the cell's previous text. Its changes are
	 * carried onto the runs like any other edit.
	 */
	transformText?: (text: string, previous?: string) => string;
}

/** The edited segments with the committed-text transform applied. */
function transformedSegments(
	element: TablePptxElement,
	row: number,
	col: number,
	segments: TextSegment[] | undefined,
	transform: TableCellEditOptions['transformText'],
): TextSegment[] | undefined {
	if (!segments || !transform) {
		return segments;
	}
	const text = segmentsPlainText(segments);
	const next = transform(text, element.tableData?.rows[row]?.cells[col]?.text);
	return next === text ? segments : editSegmentsText(segments, next);
}

/**
 * Resolve the end of an edit of cell `(row, col)` of `element`.
 */
export function resolveTableCellEditEnd(
	element: TablePptxElement,
	row: number,
	col: number,
	end: TableCellEditEnd,
	options: TableCellEditOptions = {},
): TableCellEditOutcome {
	const segments = transformedSegments(element, row, col, end.segments, options.transformText);
	let next = segments ? setTableElementCellSegments(element, row, col, segments) : element;
	const outcome: TableCellEditOutcome = {};
	switch (end.reason) {
		case 'navigate': {
			const move = resolveTableCellNavigation(next.tableData, { row, col }, end.key);
			if (move.kind === 'append-row') {
				next = appendTableElementRow(next);
				outcome.edit = { row: move.row, col: move.col, caret: move.caret };
			} else if (move.kind === 'move') {
				outcome.edit = { row: move.row, col: move.col, caret: move.caret };
			} else {
				outcome.edit = {
					row,
					col,
					caret: end.key === 'previous' || end.key === 'left' ? 'start' : 'end',
				};
			}
			break;
		}
		case 'paste':
			if (next.tableData) {
				next = fillTableCells(next, planTableGridPaste(next.tableData, { row, col }, end.grid));
			}
			outcome.select = { row, col };
			break;
		case 'exit':
			outcome.select = { row, col };
			break;
		default:
			break;
	}
	if (next !== element) {
		outcome.element = next;
	}
	return outcome;
}

/** A cell's content as the editor opens it (shared so every binding seeds it alike). */
export function tableCellEditorSegments(cell: PptxTableCell | undefined): TextSegment[] {
	return cell ? tableCellSegments(cell) : [{ text: '', style: {} }];
}

/**
 * Paste clipboard text into a table with a cell (or a range of cells)
 * selected and no cell being edited: tab-separated text fills cells from the
 * selection's top-left cell, plain text goes into that cell. Returns the
 * updated element, or `undefined` when there is nothing to paste.
 */
export function resolveTableCellsPaste(
	element: TablePptxElement,
	cells: ReadonlyArray<{ row: number; col: number }>,
	text: string,
): TablePptxElement | undefined {
	const data = element.tableData;
	if (!data || cells.length === 0 || !text) {
		return undefined;
	}
	const anchor = cells.reduce((top, cell) =>
		cell.row < top.row || (cell.row === top.row && cell.col < top.col) ? cell : top,
	);
	const grid = isTabularPaste(text) ? parseTabularText(text) : [[text.replace(/\r\n?/g, '\n')]];
	const next = fillTableCells(element, planTableGridPaste(data, anchor, grid));
	return next === element ? undefined : next;
}
