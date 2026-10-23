import { createEditorHistoryController } from 'ooxml-ui/pptx';
import type { EditorHistoryController, EditorHistoryState as HistoryFlags } from 'ooxml-ui/pptx';

import { controllerState } from '../state/controller-state';
import type { ControllerState } from '../state/controller-state';
import type { EditorSnapshot } from './editor-document-state';

/**
 * The editor's undo/redo stack: the shared history controller (one stack, one
 * default depth for every binding), read through the Svelte controller adapter
 * so `canUndo` / `canRedo` and the labels are reactive without any manual
 * re-sync after a mutation.
 */
export class EditorHistoryState {
	readonly controller: EditorHistoryController<EditorSnapshot> =
		createEditorHistoryController<EditorSnapshot>();
	readonly #state: ControllerState<HistoryFlags> = controllerState(this.controller);

	get canUndo(): boolean {
		return this.#state.current.canUndo;
	}

	get canRedo(): boolean {
		return this.#state.current.canRedo;
	}

	/** Description of the action Undo would revert, when one was recorded. */
	get undoLabel(): string | undefined {
		return this.#state.current.undoLabel;
	}

	/** Description of the action Redo would re-apply, when one was recorded. */
	get redoLabel(): string | undefined {
		return this.#state.current.redoLabel;
	}

	/** Kept for callers; the controller adapter keeps the flags current on its own. */
	sync(): void {}

	/**
	 * Apply the File > Options "maximum number of undos" value. The oldest
	 * entries beyond the new maximum are dropped; the rest stay undoable.
	 */
	setDepth(depth: number): void {
		this.controller.setMaxDepth(depth);
	}

	record(snapshot: EditorSnapshot, label = ''): void {
		this.controller.record(snapshot, label);
	}

	clear(): void {
		this.controller.clear();
	}

	undo(current: EditorSnapshot): EditorSnapshot | undefined {
		return this.controller.undo(current)?.snapshot;
	}

	redo(current: EditorSnapshot): EditorSnapshot | undefined {
		return this.controller.redo(current)?.snapshot;
	}
}
