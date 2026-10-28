import type { Doc, UndoManager } from 'yjs';

import { YDOC_ASSETS_KEY } from '../collaboration-assets';
import { YDOC_SLIDES_KEY } from '../collaboration-sync';
import type { PptxController } from './controller';
import { createControllerStore } from './controller';

export interface CollaborationHistoryState {
	readonly canUndo: boolean;
	readonly canRedo: boolean;
}

/** A room's local operations, never transactions received from another peer. */
export interface CollaborationHistory extends PptxController<CollaborationHistoryState> {
	undo(): void;
	redo(): void;
	/** Prevent the next operation from merging with the preceding one. */
	stopCapturing(): void;
	clear(): void;
	destroy(): void;
}

/**
 * The first populated room state is a baseline, not an undoable edit. The
 * origin must uniquely identify this editor's writes, including when several
 * editors share a host-owned Y.Doc.
 */
export function createCollaborationHistory(
	doc: Doc,
	UndoManagerClass: typeof UndoManager,
	localOrigin: unknown,
): CollaborationHistory {
	const slides = doc.getArray(YDOC_SLIDES_KEY);
	let established = slides.length > 0;
	const manager = new UndoManagerClass([slides, doc.getMap(YDOC_ASSETS_KEY)], {
		trackedOrigins: new Set([localOrigin]),
		captureTransaction: () => established,
	});
	const establish = (): void => {
		if (slides.length > 0) established = true;
	};
	doc.on('afterTransaction', establish);
	const store = createControllerStore<CollaborationHistoryState>({
		canUndo: false,
		canRedo: false,
	});
	const sync = (): void => {
		const next = { canUndo: manager.canUndo(), canRedo: manager.canRedo() };
		const current = store.getSnapshot();
		if (next.canUndo !== current.canUndo || next.canRedo !== current.canRedo) store.set(next);
	};
	manager.on('stack-item-added', sync);
	manager.on('stack-item-popped', sync);
	manager.on('stack-cleared', sync);
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		undo: () => {
			manager.undo();
		},
		redo: () => {
			manager.redo();
		},
		stopCapturing: () => manager.stopCapturing(),
		clear: () => manager.clear(),
		destroy: () => {
			doc.off('afterTransaction', establish);
			manager.destroy();
		},
	};
}
