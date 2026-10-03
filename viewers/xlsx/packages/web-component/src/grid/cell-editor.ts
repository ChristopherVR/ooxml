// The in-cell editor and the grid's EditDriver. Its textarea doubles as the grid's focus sink
// while not editing (typing starts Enter mode, IME composition works, copy/paste events land
// here). Commit validates with the core (data validation) before writing through the session.
import {
	cellInputText,
	formatAddress,
	parseCellInput,
	translateFormula,
	type CellAddress,
	type CellRange,
} from '@christophervr/xlsx-core';
import { showValidationAlert } from './alert.js';
import { cssFont } from './cell-paint.js';
import { h, place } from './dom.js';
import {
	editBridge,
	type BeginOptions,
	type CommitMove,
	type EditBridge,
	type EditDriver,
} from './edit-bridge.js';
import { createFormulaField, type FormulaField } from './formula-field.js';
import { closeParens, isFormulaText, referenceTarget } from './formula-text.js';
import { createFunctionAssist, type FunctionAssist } from './function-assist.js';
import type { GridView } from './grid-view.js';
import { PointMode } from './point-mode.js';

const MAX_FILL_CELLS = 20_000;

export class CellEditor implements EditDriver {
	readonly host: HTMLDivElement;
	readonly field: FormulaField;
	readonly assist: FunctionAssist;
	readonly bridge: EditBridge;
	readonly point: PointMode;
	#view: GridView;
	#afterCommit: (move: CommitMove) => void;
	#pending = false;
	#detach: () => void;

	constructor(view: GridView, afterCommit: (move: CommitMove) => void) {
		this.#view = view;
		this.#afterCommit = afterCommit;
		const { ctx, doc } = view;
		this.host = h(doc, 'div', 'xg-editor-host');
		this.field = createFormulaField(doc, 'xg-cell-field', ctx.t('Cell editor'));
		this.field.input.classList.add('xg-sink-input');
		this.host.append(this.field.root);
		view.view.append(this.host);
		this.assist = createFunctionAssist(doc, view.view, ctx.t);
		this.bridge = editBridge(ctx);
		this.point = new PointMode(view, this);
		this.#detach = this.bridge.attach(this);
		this.field.input.addEventListener('input', () => this.#onInput());
		this.field.input.addEventListener('compositionstart', () => {
			if (!this.editing) this.begin({ mode: 'enter', text: '' });
		});
		view.hooks.add(() => this.position());
	}

	get editing(): boolean {
		return this.bridge.state().editing;
	}

	focus(): void {
		const input = this.field.input;
		if (
			input.ownerDocument.activeElement !== input ||
			(input.getRootNode() as ShadowRoot).activeElement !== input
		)
			input.focus({ preventScroll: true });
	}

	#onInput(): void {
		const input = this.field.input;
		if (!this.editing) {
			const text = input.value;
			if (!text || this.#view.ctx.readOnly()) {
				input.value = '';
				return;
			}
			this.begin({ mode: 'enter', text, caret: input.selectionEnd ?? text.length });
			return;
		}
		this.point.reset();
		this.bridge.update(input.value, input.selectionEnd ?? input.value.length, 'cell');
		this.#refreshAssist();
	}

	begin(options: BeginOptions): boolean {
		const { ctx } = this.#view;
		const session = ctx.session();
		const workbook = ctx.workbook();
		if (ctx.readOnly() || !session || !workbook || this.#pending) return false;
		if (this.editing) {
			if (options.source === 'bar') this.bridge.set({ source: 'bar' }, 'grid');
			return true;
		}
		const sheet = ctx.activeSheet();
		const at = { ...ctx.selection.get().active };
		const mode = options.mode ?? 'edit';
		const text =
			options.text ?? (mode === 'enter' ? '' : cellInputText(workbook, sheet, at.row, at.col));
		const caret = Math.min(options.caret ?? text.length, text.length);
		const source = options.source ?? 'cell';
		this.host.classList.add('xg-editing');
		this.field.setText(text, caret);
		this.bridge.set({ editing: true, mode, text, caret, source, sheet, address: at }, 'grid');
		if (source === 'cell') this.focus();
		this.#refresh();
		ctx.requestRender();
		return true;
	}

	changed(source: 'cell' | 'bar'): void {
		if (source === 'bar') {
			const { text, caret } = this.bridge.state();
			this.field.setText(text, caret);
			this.point.reset();
		}
		this.#refresh();
	}

	/** Writes text and caret from the grid side (point mode, F4) and tells the formula bar. */
	setText(text: string, caret: number): void {
		this.field.setText(text, caret);
		this.bridge.update(text, caret, 'cell');
		this.bridge.set({ source: 'cell' }, 'cell');
		this.#refresh();
	}

	setMode(mode: 'enter' | 'edit' | 'point'): void {
		this.bridge.set({ mode }, 'grid');
		this.#view.ctx.requestRender();
	}

	commit(move: CommitMove, allSelected: boolean): boolean {
		const state = this.bridge.state();
		const { ctx } = this.#view;
		const session = ctx.session();
		if (!state.editing || !session || this.#pending) return false;
		const text = closeParens(state.text);
		if (!isFormulaText(text) && !allSelected) {
			const parsed = parseCellInput(text, { date1904: session.workbook.date1904 });
			const result = session.validate(
				state.sheet,
				state.address.row,
				state.address.col,
				parsed.value,
			);
			if (!result.ok) {
				this.#pending = true;
				void showValidationAlert(ctx, result.style, result.message, result.title).then((answer) => {
					this.#pending = false;
					if (answer === 'yes' || answer === 'ok') this.#apply(text, move, false);
					else if (answer === 'cancel') this.cancel();
					else this.focus();
				});
				return false;
			}
		}
		this.#apply(text, move, allSelected);
		return true;
	}

	#apply(text: string, move: CommitMove, allSelected: boolean): void {
		const state = this.bridge.state();
		const session = this.#view.ctx.session();
		this.end();
		if (!session) return;
		const { sheet, address } = state;
		try {
			if (allSelected)
				fillSelection(session, sheet, address, this.#view.ctx.selection.get().ranges, text);
			else session.setCellInput(sheet, address.row, address.col, text);
		} catch (error) {
			this.#view.ctx.toast(
				this.#view.ctx.t(error instanceof Error ? error.message : String(error)),
				'error',
			);
		}
		this.#afterCommit(move);
	}

	cancel(): void {
		if (!this.editing) return;
		this.end();
		this.focus();
	}

	/** Leaves edit mode without writing. */
	end(): void {
		this.host.classList.remove('xg-editing');
		this.field.setText('', 0);
		this.point.reset();
		this.assist.close();
		this.#view.overlay.references = undefined;
		this.bridge.set({ editing: false, mode: 'ready', text: '', caret: 0, source: 'cell' }, 'grid');
		this.focus();
		this.#view.ctx.requestRender();
		this.#view.schedule();
	}

	#refresh(): void {
		this.#updateReferences();
		this.position();
	}

	#refreshAssist(): void {
		const box = this.host.getBoundingClientRect();
		const outer = this.#view.view.getBoundingClientRect();
		this.assist.update(this.field, {
			left: box.left - outer.left,
			top: box.bottom - outer.top + 2,
		});
	}

	#updateReferences(): void {
		const view = this.#view;
		const sheet = view.sheet();
		const state = this.bridge.state();
		if (!state.editing || !sheet || state.sheet !== view.sheetIndex()) {
			view.overlay.references = undefined;
			view.schedule();
			return;
		}
		const refs: { range: CellRange; color: string }[] = [];
		for (const ref of this.field.references()) {
			const target = referenceTarget(ref.text);
			if (!target) continue;
			if (target.sheet !== undefined && target.sheet.toLowerCase() !== sheet.name.toLowerCase())
				continue;
			refs.push({ range: target.range, color: ref.color });
		}
		view.overlay.references = refs.length ? refs : undefined;
		view.schedule();
	}

	/** Keeps the editor over its cell, growing right and down with its text like Excel. */
	position(): void {
		const view = this.#view;
		const g = view.geometry;
		if (!g) return;
		const state = this.bridge.state();
		const at: CellAddress = state.editing ? state.address : view.ctx.selection.get().active;
		const range = state.editing ? { start: at, end: at } : view.activeRange();
		const box = g.rangeBox(range);
		if (!state.editing || state.sheet !== view.sheetIndex()) {
			place(this.host, Math.max(0, box.x), Math.max(0, box.y), 1, 1);
			return;
		}
		const cell = view.cellViewAt(at.row, at.col);
		const zoom = view.metrics.zoom;
		const font = cell
			? cssFont(cell.font, zoom)
			: `${Math.round(14.67 * (zoom / 100))}px "Calibri", "Carlito", sans-serif`;
		this.host.style.font = font;
		const lines = state.text.split('\n');
		const widest = Math.max(...lines.map((line) => view.measurer.measure(line, font)));
		const lineH = Math.ceil((cell?.font.sizePx ?? 14.67) * (zoom / 100) * 1.25);
		const w = Math.min(Math.max(box.w + 1, widest + 12), Math.max(box.w + 1, g.width - box.x - 2));
		const hgt = Math.max(box.h + 1, lines.length * lineH + 4);
		place(this.host, box.x - 1, box.y - 1, w, hgt);
	}

	describe(): string {
		const state = this.bridge.state();
		return formatAddress(state.address);
	}

	destroy(): void {
		this.#detach();
		this.assist.destroy();
		this.host.remove();
	}
}

/** Ctrl+Enter: the same input in every selected cell, formulas moved relative to each cell. */
function fillSelection(
	session: NonNullable<ReturnType<GridView['ctx']['session']>>,
	sheet: number,
	origin: CellAddress,
	ranges: CellRange[],
	text: string,
): void {
	session.batch('Typing', () => {
		let count = 0;
		for (const range of ranges)
			for (let r = range.start.row; r <= range.end.row; r++)
				for (let c = range.start.col; c <= range.end.col; c++) {
					if (++count > MAX_FILL_CELLS) return;
					let input = text;
					if (isFormulaText(text) && (r !== origin.row || c !== origin.col)) {
						try {
							input = `=${translateFormula(text.slice(1), r - origin.row, c - origin.col)}`;
						} catch {
							input = text;
						}
					}
					session.setCellInput(sheet, r, c, input);
				}
	});
}
