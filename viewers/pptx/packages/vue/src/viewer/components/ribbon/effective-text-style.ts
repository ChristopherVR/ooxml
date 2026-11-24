/**
 * The text style the Home tab's Font and Paragraph toggles read, shared by
 * `TextSection.vue` and `ParagraphGroup.vue`.
 */
import { tableCellsTextFormat } from 'ooxml-ui/pptx';
import { hasTextProperties } from 'pptx-viewer-core';
import type { PptxElement, TextStyle } from 'pptx-viewer-core';

import type { TableCellEditorState } from './ribbon-types';

/**
 * Returns the text style currently in effect for toolbar toggles:
 * - For text/shape/connector elements, the element's own `textStyle`.
 * - For tables, the focused cell's (or first cell's) resolved format merged
 *   with its cell style.
 * - `undefined` otherwise.
 */
export function getEffectiveTextStyle(
	element: PptxElement | null,
	tableEditorState: TableCellEditorState | null | undefined,
): Partial<TextStyle> | undefined {
	if (!element) {
		return undefined;
	}
	if (hasTextProperties(element)) {
		return element.textStyle;
	}
	if (element.type === 'table' && element.tableData) {
		const cells =
			tableEditorState && tableEditorState.elementId === element.id
				? [{ row: tableEditorState.rowIndex, col: tableEditorState.columnIndex }]
				: undefined;
		const cell = tableEditorState
			? element.tableData.rows[tableEditorState.rowIndex]?.cells[tableEditorState.columnIndex]
			: undefined;
		const format = tableCellsTextFormat(element, cells);
		return {
			...(cell?.style as Partial<TextStyle> | undefined),
			bold: format.bold,
			italic: format.italic,
			underline: format.underline,
			strikethrough: format.strikethrough,
			...(format.fontSize !== undefined ? { fontSize: format.fontSize } : {}),
			...(format.fontFamily ? { fontFamily: format.fontFamily } : {}),
			...(format.color ? { color: format.color } : {}),
			...(format.highlightColor ? { highlightColor: format.highlightColor } : {}),
			...(format.characterSpacing !== undefined
				? { characterSpacing: format.characterSpacing }
				: {}),
			...(format.textShadowColor ? { textShadowColor: format.textShadowColor } : {}),
		};
	}
	return undefined;
}
