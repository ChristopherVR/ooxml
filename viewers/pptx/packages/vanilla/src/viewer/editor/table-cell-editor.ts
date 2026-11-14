import type { TableCellCaret, TableCellTextEditorSession } from 'ooxml-ui/pptx';
import { openTableCellTextEditor, tableCellEditorSegments } from 'ooxml-ui/pptx';
import type { TablePptxElement } from 'pptx-viewer-core';

import type { EditorOps } from './editor-operations';

export interface TableCellEditorSession {
	close(commit: boolean): void;
}

/**
 * The rendered `<td>` of a table cell, looked up after the stage repaints
 * (an appended row exists only once the store change has painted).
 */
function findCell(
	doc: Document,
	elementId: string,
	row: number,
	column: number,
): HTMLTableCellElement | null {
	for (const host of Array.from(doc.querySelectorAll('[data-element-id]'))) {
		if (host.getAttribute('data-element-id') !== elementId) {
			continue;
		}
		const cell = host.querySelector<HTMLTableCellElement>(
			`td[data-row-index="${row}"][data-cell-index="${column}"]`,
		);
		if (cell) {
			return cell;
		}
	}
	return null;
}

/**
 * Edit a rendered table cell with the shared rich cell editor
 * (`openTableCellTextEditor`): runs keep their formatting, Enter adds a
 * paragraph, and Tab, Shift+Tab or an arrow at the text's edge commit and move
 * the editor to the next cell (Tab in the last cell appends a row first).
 * Escape commits and leaves the cell selected. The returned session follows
 * the editor from cell to cell.
 */
export function openTableCellEditor(options: {
	doc: Document;
	cell: HTMLTableCellElement;
	element: TablePptxElement;
	row: number;
	column: number;
	ops: EditorOps;
	caret?: TableCellCaret;
}): TableCellEditorSession {
	const { doc, ops } = options;
	const elementId = options.element.id;
	let current: TableCellTextEditorSession | undefined;
	let discard = false;
	let closed = false;

	const commitOutside = (event: PointerEvent): void => {
		const root = current?.element;
		if (root && !event.composedPath().includes(root)) {
			current?.destroy();
		}
	};
	const finish = (): void => {
		closed = true;
		current = undefined;
		doc.removeEventListener('pointerdown', commitOutside, true);
	};

	const open = (
		cell: HTMLTableCellElement,
		element: TablePptxElement,
		row: number,
		column: number,
		caret: TableCellCaret,
	): void => {
		const original = doc.createDocumentFragment();
		while (cell.firstChild) {
			original.appendChild(cell.firstChild);
		}
		current = openTableCellTextEditor({
			host: cell,
			segments: tableCellEditorSegments(
				element.tableData?.rows[row]?.cells[column],
				element,
				row,
				column,
			),
			caret,
			rtl: element.tableData?.rtl === true,
			onEnd: (end) => {
				const root = current?.element;
				current = undefined;
				root?.remove();
				if (discard) {
					cell.replaceChildren(original);
					finish();
					return;
				}
				const outcome = ops.endTableCellEdit(elementId, row, column, end);
				if (cell.isConnected && !outcome?.element) {
					cell.replaceChildren(original);
				}
				const next = outcome?.edit;
				if (!next) {
					finish();
					return;
				}
				// Move to the next cell once the store change has painted.
				const reopen = (attempt: number): void => {
					const target = findCell(doc, elementId, next.row, next.col);
					const selected = ops.selectedElement();
					const latest = selected?.id === elementId ? selected : undefined;
					if (target && !closed) {
						open(
							target,
							latest?.type === 'table' ? latest : element,
							next.row,
							next.col,
							next.caret,
						);
					} else if (attempt < 3 && !closed) {
						requestAnimationFrame(() => reopen(attempt + 1));
					} else {
						finish();
					}
				};
				queueMicrotask(() => reopen(0));
			},
		});
	};

	doc.addEventListener('pointerdown', commitOutside, true);
	open(options.cell, options.element, options.row, options.column, options.caret ?? 'all');
	return {
		close(commit: boolean): void {
			if (closed) {
				return;
			}
			discard = !commit;
			const session = current;
			if (session) {
				session.destroy();
			} else {
				finish();
			}
		},
	};
}
