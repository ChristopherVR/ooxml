// Excel's Point mode: while a formula is being typed and the caret accepts a reference, arrow
// keys or mouse clicks insert (and then keep replacing) a reference to the pointed cell or range.
import {
	formatAddress,
	formatRange,
	navigate,
	type CellAddress,
	type NavigationKey,
} from '@christophervr/xlsx-core';
import { acceptsReference, isFormulaText, splice } from './formula-text.js';
import type { GridView } from './grid-view.js';
import { spanRange } from './selection-ops.js';

export interface PointEditor {
	readonly bridge: {
		state(): { editing: boolean; text: string; caret: number; sheet: number; address: CellAddress };
	};
	setText(text: string, caret: number): void;
	setMode(mode: 'enter' | 'edit' | 'point'): void;
}

interface PointedRef {
	start: number;
	end: number;
	anchor: CellAddress;
	at: CellAddress;
}

export class PointMode {
	#ref: PointedRef | undefined;
	#view: GridView;
	#editor: PointEditor;

	constructor(view: GridView, editor: PointEditor) {
		this.#view = view;
		this.#editor = editor;
	}

	reset(): void {
		this.#ref = undefined;
	}

	get active(): boolean {
		return this.#ref !== undefined;
	}

	/** True when a pointed reference would be inserted (or replace the current one). */
	canPoint(): boolean {
		const state = this.#editor.bridge.state();
		if (!state.editing || !isFormulaText(state.text) || state.sheet !== this.#view.sheetIndex())
			return false;
		if (this.#ref && state.caret === this.#ref.end) return true;
		return acceptsReference(state.text, state.caret);
	}

	#write(anchor: CellAddress, at: CellAddress): void {
		const state = this.#editor.bridge.state();
		const same = anchor.row === at.row && anchor.col === at.col;
		const ref = same ? formatAddress(at) : formatRange(spanRange(anchor, at));
		const start = this.#ref && state.caret === this.#ref.end ? this.#ref.start : state.caret;
		const end = this.#ref && state.caret === this.#ref.end ? this.#ref.end : state.caret;
		const next = splice(state.text, start, end, ref);
		this.#ref = { start, end: next.caret, anchor, at };
		this.#editor.setText(next.text, next.caret);
		this.#editor.setMode('point');
		this.#view.reveal(at.row, at.col);
	}

	/** Arrow key in point mode; Shift extends the pointed range. */
	move(key: NavigationKey, extend: boolean): boolean {
		if (!this.canPoint()) return false;
		const sheet = this.#view.sheet();
		if (!sheet) return false;
		const state = this.#editor.bridge.state();
		const base = this.#ref?.at ?? state.address;
		const at = navigate(sheet, base, key, 20);
		const anchor = extend ? (this.#ref?.anchor ?? base) : at;
		this.#write(anchor, at);
		return true;
	}

	/** Mouse pointing: `anchor` where the press started, `at` where the pointer is. */
	pointRange(anchor: CellAddress, at: CellAddress): boolean {
		if (!this.canPoint()) return false;
		this.#write(anchor, at);
		return true;
	}
}
