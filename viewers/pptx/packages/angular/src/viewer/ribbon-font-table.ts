/**
 * ribbon-font-table.ts: the Home tab's Font group on a selected table. Every
 * run of the selected cells (the Shift+click range, else the active cell; the
 * whole table when no cell of it is selected) takes the format, through the
 * shared `tableCellsTextStylePatch`; the toggles read the first selected cell.
 * Pure (no DI), so it is testable without TestBed.
 */
import type {
	ChangeCaseMode,
	TableCellRef,
	TableCellsFontUpdate,
	TableCellsTextFormat,
} from 'ooxml-ui/pptx';
import {
	stepFontSizePt,
	tableCellsChangeCasePatch,
	tableCellsTextFormat,
	tableCellsTextStylePatch,
	tableTextShadowUpdate,
} from 'ooxml-ui/pptx';
import type { PptxElement, TablePptxElement, TextStyle } from 'pptx-viewer-core';

import { fontHomeAction } from './ribbon-font-home';
import type { TableCellSelection } from './table-selection.service';

/** The selected cells of `element`, or `undefined` for the whole table. */
export function tableFontCells(
	element: TablePptxElement,
	selection: TableCellSelection | null | undefined,
): TableCellRef[] | undefined {
	if (!selection || selection.elementId !== element.id) {
		return undefined;
	}
	return selection.selectedCells?.length
		? selection.selectedCells
		: [{ row: selection.rowIndex, col: selection.columnIndex }];
}

/** The ribbon's text style for a selected table (sizes in px, as text boxes report them). */
export function tableFontStyle(
	element: PptxElement | null,
	selection: TableCellSelection | null | undefined,
): TextStyle | null {
	if (element?.type !== 'table') {
		return null;
	}
	const format = tableCellsTextFormat(element, tableFontCells(element, selection));
	return {
		bold: format.bold,
		italic: format.italic,
		underline: format.underline,
		strikethrough: format.strikethrough,
		...(format.fontSize === undefined ? {} : { fontSize: (format.fontSize * 96) / 72 }),
		...(format.fontFamily ? { fontFamily: format.fontFamily } : {}),
		...(format.color ? { color: format.color } : {}),
		...(format.highlightColor ? { highlightColor: format.highlightColor } : {}),
		...(format.characterSpacing !== undefined ? { characterSpacing: format.characterSpacing } : {}),
		...(format.textShadowColor ? { textShadowColor: format.textShadowColor } : {}),
	};
}

function tableUpdate(
	id: string,
	value: unknown,
	current: TableCellsTextFormat,
): TableCellsFontUpdate | undefined {
	const size = (pt: number) => ({ fontSize: Math.min(400, Math.max(1, pt)) });
	const action = fontHomeAction(id as Parameters<typeof fontHomeAction>[0]);
	if (action?.kind === 'toggle') {
		return { [action.flag]: !current[action.flag] };
	}
	if (action?.kind === 'shadow') {
		return tableTextShadowUpdate(Boolean(current.textShadowColor));
	}
	if (action?.kind === 'step') {
		const step = action.direction === 1 ? 'increase' : 'decrease';
		return size(stepFontSizePt(current.fontSize ?? 18, step));
	}
	if (action?.kind === 'clear') {
		return {
			bold: false,
			italic: false,
			underline: false,
			strikethrough: false,
			highlightColor: undefined,
			characterSpacing: 0,
			...tableTextShadowUpdate(true),
		};
	}
	switch (id) {
		case 'home.font.fontColor':
			return { color: String(value) };
		case 'home.font.highlightColor':
			return { highlightColor: String(value) };
		case 'home.font.characterSpacing':
			return { characterSpacing: Number(value) };
		case 'home.font.fontFamily':
			return { fontFamily: String(value) };
		case 'home.font.fontSize':
			return Number(value) > 0 ? size(Number(value)) : undefined;
		default:
			return undefined;
	}
}

/**
 * The element patch a Font group request makes on a selected table, or
 * `undefined` when the element is not a table or the request does not apply
 * to table cells (the caller then handles it as for a text box).
 */
export function tableFontRequestPatch(
	element: PptxElement | null,
	selection: TableCellSelection | null | undefined,
	id: string,
	value: unknown,
): Partial<PptxElement> | undefined {
	if (element?.type !== 'table') {
		return undefined;
	}
	const cells = tableFontCells(element, selection);
	if (id === 'home.font.changeCase') {
		return tableCellsChangeCasePatch(element, cells, value as ChangeCaseMode);
	}
	const update = tableUpdate(id, value, tableCellsTextFormat(element, cells));
	return update ? tableCellsTextStylePatch(element, cells, update) : undefined;
}
