import type { ChangeCaseMode, TableCellsFontUpdate, TableCellsTextFormat } from 'ooxml-ui/pptx';
import {
	tableCellsChangeCasePatch,
	tableCellsTextFormat,
	tableCellsTextStylePatch,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';

/** The selected cells of the selected table (the Shift range, else the active cell). */
function selectedCells(state: ViewerState): Array<{ row: number; col: number }> | undefined {
	if (state.selectedTableCells.length > 0) {
		return state.selectedTableCells.map(({ row, column }) => ({ row, col: column }));
	}
	return state.selectedTableCell
		? [{ row: state.selectedTableCell.row, col: state.selectedTableCell.column }]
		: undefined;
}

/**
 * The Home tab's character formatting for a selected table: applied to every
 * run of the selected cells (the whole table when no cell is selected) through
 * the shared `tableCellsTextStylePatch`. `undefined` for any other element, so
 * the caller falls back to its text-box formatting.
 */
export function tableTextFormatPatch(
	store: Store<ViewerState>,
	el: PptxElement,
	build: (current: TableCellsTextFormat) => TableCellsFontUpdate,
): Partial<PptxElement> | undefined {
	if (el.type !== 'table') {
		return undefined;
	}
	const cells = selectedCells(store.get());
	return tableCellsTextStylePatch(el, cells, build(tableCellsTextFormat(el, cells)));
}

/** Rewrite the character case of every run in the selected cells of a table. */
export function tableTextChangeCasePatch(
	store: Store<ViewerState>,
	el: PptxElement,
	mode: ChangeCaseMode,
): Partial<PptxElement> | undefined {
	if (el.type !== 'table') {
		return undefined;
	}
	return tableCellsChangeCasePatch(el, selectedCells(store.get()), mode);
}
