/**
 * table-render.tsx: renderTableElement
 *
 * Every table renders from the structured `tableData` the core parses (the same
 * model the other four bindings paint), through the shared `tableCellCss` and
 * `cellParagraphBlocks` descriptors. React keeps only the JSX mapping:
 *   - table-render-data.tsx       - renderTableFromTableData
 *   - table-render-cell-input.tsx - TableCellInput inline editor
 *   - table-render-resize.tsx     - TableResizeOverlay drag overlay
 */
import { translationsEn } from 'ooxml-ui/pptx/i18n';
import type { PptxElement, TablePptxElement } from 'pptx-viewer-core';
import React from 'react';

import type { TableCellEditorState } from '../types';
import type { TableStyleContext } from './table-band-style';
import { renderTableFromTableData } from './table-render-data';

export function renderTableElement(
	element: PptxElement,
	textStyle: React.CSSProperties,
	options?: {
		editable?: boolean;
		selectedCell?: TableCellEditorState | null;
		onSelectCell?: (cell: TableCellEditorState) => void;
		onCommitCellEdit?: (rowIndex: number, colIndex: number, text: string) => void;
		onResizeColumns?: (newWidths: number[]) => void;
		onResizeRow?: (rowIndex: number, newHeight: number) => void;
		styleCtx?: TableStyleContext;
	},
): React.ReactNode {
	const tableEl = element as TablePptxElement;
	if (tableEl.tableData && tableEl.tableData.rows.length > 0) {
		return renderTableFromTableData(tableEl, textStyle, options);
	}
	return (
		<div className='w-full h-full flex items-center justify-center text-[11px] text-white/80 pointer-events-none'>
			{translationsEn['pptx.table.title']}
		</div>
	);
}
