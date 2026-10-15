// The in-cell text box of Page Layout view. It sits over the cell on its page, commits through the
// edit session (so undo, recalculation and collaboration behave as in Normal view) and moves the
// selection like Excel: Enter down, Tab right, Escape discards.
import type { PageViewRect } from 'ooxml-core/xlsx';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { selectCell } from 'ooxml-core/xlsx/ui';

export interface PageEditorOptions {
	ctx: EditorContext;
	doc: Document;
	scaled: HTMLElement;
	/** The cell's rectangle in the scaled canvas, or undefined when no page shows it. */
	rectOf(row: number, col: number): PageViewRect | undefined;
	scale(): number;
	done(): void;
}

export interface PageEditor {
	readonly editing: boolean;
	begin(row: number, col: number, text: string): void;
	cancel(): void;
	rectFor(at: { row: number; col: number }): PageViewRect | undefined;
}

export function createPageEditor(o: PageEditorOptions): PageEditor {
	let input: HTMLInputElement | undefined;
	let endCurrent: ((next: 'down' | 'right' | 'none', save: boolean) => void) | undefined;
	const close = () => {
		const node = input;
		input = undefined;
		node?.remove();
	};
	const commit = (row: number, col: number, text: string, next: 'down' | 'right' | 'none') => {
		const session = o.ctx.session();
		const sheetIndex = o.ctx.activeSheet();
		close();
		if (session) {
			try {
				session.setCellInput(sheetIndex, row, col, text);
			} catch (error) {
				o.ctx.toast(error instanceof Error ? error.message : String(error), 'error');
			}
		}
		if (next !== 'none') {
			const sheet = o.ctx.workbook()?.sheets[sheetIndex];
			o.ctx.selection.set(
				selectCell(sheetIndex, sheet, {
					row: row + (next === 'down' ? 1 : 0),
					col: col + (next === 'right' ? 1 : 0),
				}),
			);
		}
		o.done();
	};
	return {
		get editing() {
			return input !== undefined;
		},
		rectFor: (at) => o.rectOf(at.row, at.col),
		begin(row, col, text) {
			if (o.ctx.readOnly() || !o.ctx.session()) return;
			const rect = o.rectOf(row, col);
			if (!rect) return;
			endCurrent?.('none', false);
			const node = o.doc.createElement('input');
			node.className = 'xpl-edit';
			node.type = 'text';
			node.value = text;
			node.spellcheck = false;
			node.setAttribute('aria-label', o.ctx.t('Edit cell'));
			node.style.cssText = `left:${rect.x}px;top:${rect.y}px;min-width:${rect.width}px;width:${Math.max(rect.width, 80)}px;height:${rect.height}px;font-size:${14.67 * o.scale()}px`;
			let finished = false;
			const end = (next: 'down' | 'right' | 'none', save: boolean) => {
				if (finished) return;
				finished = true;
				endCurrent = undefined;
				if (save) commit(row, col, node.value, next);
				else {
					close();
					o.done();
				}
			};
			node.addEventListener('keydown', (event) => {
				event.stopPropagation();
				if (event.key === 'Enter') {
					event.preventDefault();
					end('down', true);
				} else if (event.key === 'Tab') {
					event.preventDefault();
					end('right', true);
				} else if (event.key === 'Escape') {
					event.preventDefault();
					end('none', false);
				}
			});
			node.addEventListener('blur', () => end('none', true));
			input = node;
			endCurrent = end;
			o.scaled.append(node);
			node.focus();
			node.setSelectionRange(node.value.length, node.value.length);
		},
		cancel() {
			endCurrent?.('none', false);
		},
	};
}
