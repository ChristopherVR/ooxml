/**
 * table-cell-navigation.ts: where the caret goes when it leaves a table cell
 * being edited, decided once for every binding.
 *
 * PowerPoint's table text keys: Tab moves to the next cell and selects its
 * text, Shift+Tab to the previous one; Tab in the last cell adds a row and
 * moves into it. An arrow key at the edge of the cell's text moves to the
 * neighbouring cell (Left at the start, Right at the end, Up on the first
 * line, Down on the last). Merged continuation cells are skipped: the move
 * lands on the cell that owns the merge. `left` and `right` are logical
 * (back and forward in reading order); in right-to-left text the arrow keys
 * map onto them mirrored.
 *
 * @module render/table-cell-navigation
 */
import type { PptxTableData } from 'ooxml-core/pptx';

import { mapInlineTextFormatKey } from './editor-keymap-chords';
import type { InlineTextFormatProperty } from './editor-keymap-chords';

/** A move out of the edited cell. */
export type TableCellNavigationKey = 'next' | 'previous' | 'up' | 'down' | 'left' | 'right';

/** Where the caret lands in the cell it moves to. */
export type TableCellCaret = 'all' | 'start' | 'end';

/** The decision for one move. */
export type TableCellNavigation =
	| { kind: 'move'; row: number; col: number; caret: TableCellCaret }
	/** Tab in the last cell: append a row, then edit `(row, col)` of it. */
	| { kind: 'append-row'; row: number; col: number; caret: 'all' }
	/** At the table's edge: the caret stays where it is. */
	| { kind: 'stay' };

interface CellPos {
	row: number;
	col: number;
}

/** The cell that owns grid position `(row, col)`, walking back over merge continuations. */
export function tableCellOwner(data: PptxTableData, row: number, col: number): CellPos | undefined {
	let r = row;
	let c = col;
	for (let guard = 0; guard < 10_000; guard++) {
		const cell = data.rows[r]?.cells[c];
		if (!cell) {
			return undefined;
		}
		if (cell.hMerge && c > 0) {
			c--;
		} else if (cell.vMerge && r > 0) {
			r--;
		} else {
			return { row: r, col: c };
		}
	}
	return undefined;
}

/** Every owning (non-continuation) cell, in reading order. */
function anchorCells(data: PptxTableData): CellPos[] {
	const cells: CellPos[] = [];
	data.rows.forEach((row, r) =>
		row.cells.forEach((cell, c) => {
			if (!cell.hMerge && !cell.vMerge) {
				cells.push({ row: r, col: c });
			}
		}),
	);
	return cells;
}

/**
 * Resolve a move from the edited cell `from`. `left` and `right` follow the
 * reading order (wrapping between rows) like Shift+Tab and Tab, but put the
 * caret at the end or the start of the cell instead of selecting its text.
 */
export function resolveTableCellNavigation(
	data: PptxTableData | undefined,
	from: CellPos,
	key: TableCellNavigationKey,
): TableCellNavigation {
	if (!data || data.rows.length === 0) {
		return { kind: 'stay' };
	}
	const start = tableCellOwner(data, from.row, from.col) ?? from;
	const step = key;
	if (step === 'up' || step === 'down') {
		const span = data.rows[start.row]?.cells[start.col]?.rowSpan ?? 1;
		const row = step === 'up' ? start.row - 1 : start.row + Math.max(1, span);
		const target = row >= 0 ? tableCellOwner(data, row, start.col) : undefined;
		return target
			? { kind: 'move', ...target, caret: step === 'up' ? 'end' : 'start' }
			: { kind: 'stay' };
	}
	const order = anchorCells(data);
	const index = order.findIndex((cell) => cell.row === start.row && cell.col === start.col);
	const forward = step === 'next' || step === 'right';
	const target = index < 0 ? undefined : order[index + (forward ? 1 : -1)];
	if (target) {
		const caret: TableCellCaret =
			step === 'next' || step === 'previous' ? 'all' : forward ? 'start' : 'end';
		return { kind: 'move', ...target, caret };
	}
	if (step === 'next') {
		const firstCol = data.rows[data.rows.length - 1]?.cells.findIndex((cell) => !cell.hMerge) ?? 0;
		return { kind: 'append-row', row: data.rows.length, col: Math.max(0, firstCol), caret: 'all' };
	}
	return { kind: 'stay' };
}

/** Where the caret is in the cell's text, for the arrow-key boundaries. */
export interface TableCellCaretState {
	/** The selection is a caret (nothing selected). */
	collapsed: boolean;
	atStart: boolean;
	atEnd: boolean;
	onFirstLine: boolean;
	onLastLine: boolean;
	/** The cell's text runs right to left (a right-to-left table or paragraph). */
	rtl?: boolean;
}

/** What a key press in the cell editor does. */
export type TableCellEditorKeyAction =
	| { kind: 'navigate'; key: TableCellNavigationKey }
	| { kind: 'exit' }
	| { kind: 'paragraph' }
	| { kind: 'line-break' }
	| { kind: 'tab-character' }
	| { kind: 'format'; property: InlineTextFormatProperty }
	| { kind: 'suppress' };

/** The parts of a keyboard event the mapping reads. */
export interface TableCellEditorKeyInput {
	key: string;
	shiftKey: boolean;
	ctrlKey: boolean;
	metaKey: boolean;
	altKey: boolean;
	isComposing?: boolean;
}

/**
 * Map a key press in the cell editor, or `null` to let the browser edit the
 * text as usual. Escape leaves editing with the cell selected; the text typed
 * so far is kept, as in PowerPoint.
 */
export function mapTableCellEditorKey(
	input: TableCellEditorKeyInput,
	caret: TableCellCaretState,
): TableCellEditorKeyAction | null {
	if (input.isComposing) {
		return null;
	}
	const command = input.ctrlKey || input.metaKey;
	if (input.key === 'Escape') {
		return { kind: 'exit' };
	}
	if (input.key === 'Tab' && !input.altKey) {
		if (command) {
			return { kind: 'tab-character' };
		}
		return { kind: 'navigate', key: input.shiftKey ? 'previous' : 'next' };
	}
	if (input.key === 'Enter' && !input.altKey) {
		if (command) {
			return { kind: 'exit' };
		}
		return { kind: input.shiftKey ? 'line-break' : 'paragraph' };
	}
	const property = mapInlineTextFormatKey(input);
	if (property) {
		return { kind: 'format', property };
	}
	if (command && !input.altKey && input.key.toLowerCase() === 'd') {
		return { kind: 'suppress' };
	}
	if (input.shiftKey || input.altKey || command || !caret.collapsed) {
		return null;
	}
	// In right-to-left text the arrows are visual: Left walks towards the end.
	const back = caret.rtl ? 'ArrowRight' : 'ArrowLeft';
	const forward = caret.rtl ? 'ArrowLeft' : 'ArrowRight';
	switch (input.key) {
		case back:
			return caret.atStart ? { kind: 'navigate', key: 'left' } : null;
		case forward:
			return caret.atEnd ? { kind: 'navigate', key: 'right' } : null;
		case 'ArrowUp':
			return caret.onFirstLine ? { kind: 'navigate', key: 'up' } : null;
		case 'ArrowDown':
			return caret.onLastLine ? { kind: 'navigate', key: 'down' } : null;
		default:
			return null;
	}
}
