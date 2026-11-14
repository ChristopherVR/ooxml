/**
 * table-cells-paste-listener.ts: the paste entry point for a table with cells
 * selected and no cell being edited. It listens for the document's `paste`
 * event (which the browser fires on Ctrl/Cmd+V when the editor's own keymap
 * leaves the chord to it) in the capture phase, and takes the event only when
 * the binding reports a selected table cell range and the clipboard holds
 * plain text. Everything else (a text box being edited, a cell editor, which
 * handles its own paste, or no table selection) is left to the other paste
 * handlers.
 *
 * @module render/table-cells-paste-listener
 */
import type { TablePptxElement } from 'ooxml-core/pptx';

import { resolveTableCellsPaste } from './table-cell-edit-outcome';

/** The selected table and cells a paste would fill, or `undefined`. */
export interface TableCellsPasteTarget {
	element: TablePptxElement;
	cells: ReadonlyArray<{ row: number; col: number }>;
}

/**
 * Listen for pastes into a selected cell range. `target` is read at paste
 * time; `apply` receives the updated element. Returns the detach function.
 */
export function attachTableCellsPaste(
	doc: Document,
	target: () => TableCellsPasteTarget | undefined,
	apply: (element: TablePptxElement) => void,
): () => void {
	const onPaste = (event: ClipboardEvent): void => {
		const focused = event.target instanceof Element ? event.target : null;
		if (
			focused?.closest('[data-inline-editor], input, textarea, [contenteditable="true"]') ||
			event.defaultPrevented
		) {
			return;
		}
		const selected = target();
		const text = event.clipboardData?.getData('text/plain') ?? '';
		if (!selected || !text) {
			return;
		}
		const next = resolveTableCellsPaste(selected.element, selected.cells, text);
		if (!next) {
			return;
		}
		event.preventDefault();
		event.stopImmediatePropagation();
		apply(next);
	};
	doc.addEventListener('paste', onPaste, true);
	return () => doc.removeEventListener('paste', onPaste, true);
}
