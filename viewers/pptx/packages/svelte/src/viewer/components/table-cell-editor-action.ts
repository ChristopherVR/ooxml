import type { TableCellCaret, TableCellEditEnd } from 'ooxml-ui/pptx';
import { openTableCellTextEditor, tableCellEditorSegments } from 'ooxml-ui/pptx';
import type { PptxTableCell, TablePptxElement } from 'pptx-viewer-core';

export interface TableCellEditorActionOptions {
	cell: PptxTableCell | undefined;
	/** The table and the cell's position, so the editor keeps each run's markup. */
	element?: TablePptxElement;
	row?: number;
	col?: number;
	caret?: TableCellCaret;
	rtl?: boolean;
	onend: (end: TableCellEditEnd) => void;
}

/**
 * `use:tableCellEditor`: mount the shared rich cell editor
 * (`openTableCellTextEditor`) into the host node for as long as the node lives.
 * The seed is read once; destroying an editor that is still open reports a
 * `detach` end, which commits.
 */
export function tableCellEditor(
	node: HTMLElement,
	options: TableCellEditorActionOptions,
): { update(next: TableCellEditorActionOptions): void; destroy(): void } {
	let current = options;
	const session = openTableCellTextEditor({
		host: node,
		segments: tableCellEditorSegments(options.cell, options.element, options.row, options.col),
		...(options.caret ? { caret: options.caret } : {}),
		rtl: options.rtl === true,
		onEnd: (end) => current.onend(end),
	});
	return {
		update(next: TableCellEditorActionOptions) {
			current = next;
		},
		destroy() {
			session.destroy();
		},
	};
}
