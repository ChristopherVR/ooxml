// The shared in-cell / formula-bar edit state. The grid owns editing (it validates and commits);
// the formula bar reads the same text and caret and forwards what the user types. One bridge per
// editor context, created on first use so mount order does not matter.
import type { CellAddress } from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';

/** Excel's status-bar modes: Ready, Enter (typing replaces), Edit (F2, arrows move the caret), Point. */
export type EditMode = 'ready' | 'enter' | 'edit' | 'point';
export type EditSource = 'cell' | 'bar';
export type CommitMove = 'down' | 'up' | 'right' | 'left' | 'none';

export interface EditState {
	editing: boolean;
	mode: EditMode;
	text: string;
	/** Caret (selection end) in `text`. */
	caret: number;
	/** Where the user is typing. */
	source: EditSource;
	sheet: number;
	address: CellAddress;
}

export interface BeginOptions {
	text?: string;
	source?: EditSource;
	/** 'enter' replaces the content (typing), 'edit' keeps it (F2, formula bar click). */
	mode?: 'enter' | 'edit';
	caret?: number;
}

/** Implemented by the grid's editor controller. */
export interface EditDriver {
	begin(options: BeginOptions): boolean;
	commit(move: CommitMove, allSelected: boolean): boolean;
	cancel(): void;
	/** Called after the text changed from `source` (the other view must follow). */
	changed(source: EditSource): void;
}

export interface EditBridge {
	state(): Readonly<EditState>;
	onChange(listener: (state: Readonly<EditState>, origin: EditSource | 'grid') => void): () => void;
	begin(options?: BeginOptions): boolean;
	update(text: string, caret: number, source: EditSource): void;
	commit(move?: CommitMove, allSelected?: boolean): boolean;
	cancel(): void;
	/** Grid side: replaces the whole state and notifies. */
	set(next: Partial<EditState>, origin: EditSource | 'grid'): void;
	attach(driver: EditDriver): () => void;
}

const bridges = new WeakMap<EditorContext, EditBridge>();

export function createEditBridge(): EditBridge {
	let state: EditState = {
		editing: false,
		mode: 'ready',
		text: '',
		caret: 0,
		source: 'cell',
		sheet: 0,
		address: { row: 0, col: 0 },
	};
	let driver: EditDriver | undefined;
	const listeners = new Set<(s: Readonly<EditState>, origin: EditSource | 'grid') => void>();
	const emit = (origin: EditSource | 'grid') => {
		for (const listener of [...listeners]) listener(state, origin);
	};
	return {
		state: () => state,
		onChange(listener) {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
		begin: (options = {}) => driver?.begin(options) ?? false,
		update(text, caret, source) {
			if (!state.editing) return;
			state = { ...state, text, caret, source };
			driver?.changed(source);
			emit(source);
		},
		commit: (move = 'none', allSelected = false) => driver?.commit(move, allSelected) ?? false,
		cancel: () => driver?.cancel(),
		set(next, origin) {
			state = { ...state, ...next };
			emit(origin);
		},
		attach(next) {
			driver = next;
			return () => {
				if (driver === next) driver = undefined;
			};
		},
	};
}

/** The bridge of an editor context (created on first use). */
export function editBridge(ctx: EditorContext): EditBridge {
	let bridge = bridges.get(ctx);
	if (!bridge) {
		bridge = createEditBridge();
		bridges.set(ctx, bridge);
	}
	return bridge;
}
