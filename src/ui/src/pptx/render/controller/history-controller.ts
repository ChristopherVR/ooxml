/**
 * The undo/redo history controller every binding uses.
 *
 * Before this module each binding wrapped the core `EditorHistory` stack its
 * own way (React kept a second, hand-written stack), and the default depth had
 * drifted to 100 in two bindings and 120 in two others. This controller is the
 * one place that owns the stack, its default depth
 * ({@link DEFAULT_EDITOR_HISTORY_DEPTH}) and the observable flags the chrome
 * reads (`canUndo`, `canRedo`, the two labels), following the controller
 * convention in `./controller`.
 *
 * Snapshots follow the core's pre-mutation model: `record` takes the state
 * from BEFORE the edit, and the caller passes already-cloned values.
 *
 * @module render/controller/history-controller
 */

import {
	DEFAULT_EDITOR_HISTORY_DEPTH,
	EditorHistory,
} from 'ooxml-core/pptx/editor/render/editor-history';
import type { UndoRedoResult } from 'ooxml-core/pptx/editor/render/editor-history';

import type { CollaborationHistory } from './collaboration-history';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

/** What the chrome reads from the history (Undo/Redo buttons, tooltips). */
export interface EditorHistoryState {
	readonly canUndo: boolean;
	readonly canRedo: boolean;
	/** The action Undo would revert; undefined when none or unlabelled. */
	readonly undoLabel: string | undefined;
	/** The action Redo would re-apply; undefined when none or unlabelled. */
	readonly redoLabel: string | undefined;
	/** Entries on the undo stack. */
	readonly depth: number;
	/** The maximum number of undo entries kept. */
	readonly maxDepth: number;
	/** Counts committed edits: every record, and every undo or redo that applied. */
	readonly revision: number;
}

/** Options for {@link createEditorHistoryController}. */
export interface EditorHistoryControllerOptions {
	/** Maximum undo depth; defaults to {@link DEFAULT_EDITOR_HISTORY_DEPTH}. */
	maxDepth?: number;
}

/** The shared undo/redo stack plus its observable state. */
export interface EditorHistoryController<T> extends PptxController<EditorHistoryState> {
	/** Push the pre-edit `snapshot` and clear the redo stack. */
	record(snapshot: T, label?: string): void;
	/** Step back: stores `current` for redo and returns the entry to restore. */
	undo(current: T): UndoRedoResult<T> | undefined;
	/** Step forward: stores `current` for undo and returns the entry to restore. */
	redo(current: T): UndoRedoResult<T> | undefined;
	/** Drop both stacks (a new document was loaded). */
	clear(): void;
	/**
	 * Drop every undo and redo step that would bring back a state `reaches` matches, with the
	 * steps behind it (the core `EditorHistory.forget`): after a change the stack cannot reverse,
	 * such as a Slide Master view command that removed a layout from the package.
	 */
	forget(reaches: (snapshot: T) => boolean): void;
	/** Apply File > Options "Maximum number of undos"; trims the oldest entries. */
	setMaxDepth(maxDepth: number): void;
	/**
	 * Route Undo and Redo through a room's own history while one is attached (`null` detaches).
	 * The room history reverts only this editor's operations in the shared document, so a
	 * participant's later edit survives. While attached, `undo` and `redo` return `undefined`
	 * (nothing to restore: the document observer applies the result) and `record` is ignored.
	 * Returns the detach function.
	 */
	attachCollaboration(history: CollaborationHistory | null): () => void;
}

function clampDepth(depth: number | undefined): number {
	const value = depth ?? DEFAULT_EDITOR_HISTORY_DEPTH;
	return Number.isFinite(value) ? Math.max(1, Math.floor(value)) : DEFAULT_EDITOR_HISTORY_DEPTH;
}

/** Build a history controller over the core `EditorHistory` stack. */
export function createEditorHistoryController<T>(
	options: EditorHistoryControllerOptions = {},
): EditorHistoryController<T> {
	let maxDepth = clampDepth(options.maxDepth);
	const stack = new EditorHistory<T>({ maxDepth });
	let revision = 0;

	let room: CollaborationHistory | null = null;
	let unsubscribeRoom: (() => void) | null = null;

	const read = (): EditorHistoryState => ({
		canUndo: room ? room.getSnapshot().canUndo : stack.canUndo,
		canRedo: room ? room.getSnapshot().canRedo : stack.canRedo,
		undoLabel: room ? undefined : stack.undoLabel || undefined,
		redoLabel: room ? undefined : stack.redoLabel || undefined,
		depth: stack.depth,
		maxDepth,
		revision,
	});
	const store = createControllerStore<EditorHistoryState>(read());
	const sync = (): void => {
		const next = read();
		const current = store.getSnapshot();
		const same = (Object.keys(next) as (keyof EditorHistoryState)[]).every(
			(key) => next[key] === current[key],
		);
		if (!same) {
			store.set(next);
		}
	};

	const step = (result: UndoRedoResult<T> | undefined): UndoRedoResult<T> | undefined => {
		if (result) {
			revision += 1;
			sync();
		}
		return result;
	};

	const stepRoom = (available: boolean, apply: () => void): undefined => {
		if (available) {
			apply();
			revision += 1;
			sync();
		}
		return undefined;
	};

	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		record(snapshot, label = '') {
			if (room) {
				return;
			}
			stack.record(snapshot, label);
			revision += 1;
			sync();
		},
		undo(current) {
			if (room) {
				return stepRoom(room.getSnapshot().canUndo, () => room?.undo());
			}
			return step(stack.undo(current));
		},
		redo(current) {
			if (room) {
				return stepRoom(room.getSnapshot().canRedo, () => room?.redo());
			}
			return step(stack.redo(current));
		},
		clear() {
			stack.clear();
			room?.clear();
			sync();
		},
		attachCollaboration(next) {
			unsubscribeRoom?.();
			unsubscribeRoom = null;
			stack.clear();
			room = next;
			if (next) {
				unsubscribeRoom = next.subscribe(sync);
			}
			sync();
			return () => {
				if (room === next) {
					unsubscribeRoom?.();
					unsubscribeRoom = null;
					room = null;
					sync();
				}
			};
		},
		forget(reaches) {
			stack.forget(reaches);
			sync();
		},
		setMaxDepth(next) {
			maxDepth = clampDepth(next);
			stack.setMaxDepth(maxDepth);
			sync();
		},
	};
}
