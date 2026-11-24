import type { ChangeCaseMode, TableCellsFontUpdate, TableCellsTextFormat } from 'ooxml-ui/pptx';
import {
	tableCellsChangeCasePatch,
	tableCellsTextFormat,
	tableCellsTextStylePatch,
} from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { EditorState } from './editor-state.svelte';

/**
 * The Home tab's character formatting for a selected table: every run of the
 * selected cell range (the whole table when no cell is selected) takes it,
 * through the shared `tableCellsTextStylePatch`. `undefined` for any other
 * element, so the caller keeps its text-box formatting.
 */
export function tableTextFormatPatch(
	editor: EditorState,
	element: PptxElement,
	build: (current: TableCellsTextFormat) => TableCellsFontUpdate,
): Partial<PptxElement> | undefined {
	if (element.type !== 'table') {
		return undefined;
	}
	const cells = editor.tableCells.cellsFor(element.id);
	const range = cells.length > 0 ? cells : undefined;
	return tableCellsTextStylePatch(element, range, build(tableCellsTextFormat(element, range)));
}

/** Rewrite the character case of every run in the selected cells of a table. */
export function tableTextChangeCasePatch(
	editor: EditorState,
	element: PptxElement,
	mode: ChangeCaseMode,
): Partial<PptxElement> | undefined {
	if (element.type !== 'table') {
		return undefined;
	}
	const cells = editor.tableCells.cellsFor(element.id);
	return tableCellsChangeCasePatch(element, cells.length > 0 ? cells : undefined, mode);
}

/** The ribbon's toggle state for a selected table (its first selected cell). */
export function tableTextFormatOf(
	editor: EditorState,
	element: PptxElement | undefined,
): TableCellsTextFormat | undefined {
	if (element?.type !== 'table') {
		return undefined;
	}
	const cells = editor.tableCells.cellsFor(element.id);
	return tableCellsTextFormat(element, cells.length > 0 ? cells : undefined);
}
