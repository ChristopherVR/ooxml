import { attachTableCellsPaste } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { useEffect, useRef } from 'react';

import type { TableCellEditorState } from '../types';

/**
 * Paste plain text (tab-separated values fill cells from the top-left
 * selected cell, growing the table) while table cells are selected and none
 * is being edited. A cell being edited handles its own paste.
 */
export function useTableCellsPaste(input: {
	selectedElement: PptxElement | undefined | null;
	tableEditorState: TableCellEditorState | null;
	updateElementById: (id: string, updates: Partial<PptxElement>) => void;
	markDirty: () => void;
}): void {
	const latest = useRef(input);
	latest.current = input;
	useEffect(() => {
		if (typeof document === 'undefined') {
			return;
		}
		return attachTableCellsPaste(
			document,
			() => {
				const { selectedElement: element, tableEditorState: ts } = latest.current;
				if (element?.type !== 'table' || !ts || ts.isEditing) {
					return undefined;
				}
				const cells = ts.selectedCells?.length
					? ts.selectedCells
					: [{ row: ts.rowIndex, col: ts.columnIndex }];
				return { element, cells };
			},
			(next) => {
				latest.current.updateElementById(next.id, {
					tableData: next.tableData,
					...(next.rawXml ? { rawXml: next.rawXml } : {}),
				});
				latest.current.markDirty();
			},
		);
	}, []);
}
