import type { Doc, UndoManager } from 'yjs';

import { YDOC_ASSETS_KEY } from '../collaboration-assets';
import { LOCAL_SYNC_ORIGIN } from '../collaboration-reconcile';
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

/** What decides whether local edits are being written into the room. */
export interface CollaborationHistoryReadiness {
	hasDoc: boolean;
	/** A host-owned document stays editable while offline. */
	external: boolean;
	connected: boolean;
	synced: boolean;
	viewer: boolean;
}

/**
 * Undo goes through the room only while this editor's edits reach it: otherwise an offline or
 * read-only edit would be missing from the room history and could not be undone.
 */
export function collaborationHistoryReady(state: CollaborationHistoryReadiness): boolean {
	return state.hasDoc && !state.viewer && (state.external || (state.connected && state.synced));
}

/** Where a room's history plugs in: the editor history controller. */
export interface CollaborationHistoryTarget {
	attachCollaboration(history: CollaborationHistory | null): () => void;
}

/**
 * Give `target` a history over `doc` for as long as the returned release has not run. `loadYjs`
 * must resolve the same Yjs instance that built `doc`; a build without `UndoManager` leaves the
 * local stack in charge. Both an owned room and a host-owned document go through this, so every
 * binding wires collaboration undo with the one call.
 */
export function bindCollaborationHistory(
	target: CollaborationHistoryTarget,
	doc: unknown,
	loadYjs: () => Promise<unknown>,
	localOrigin: unknown = LOCAL_SYNC_ORIGIN,
): () => void {
	let released = false;
	let detach: (() => void) | null = null;
	let history: CollaborationHistory | null = null;
	void loadYjs().then((Y) => {
		const Manager = (Y as { UndoManager?: typeof UndoManager }).UndoManager;
		if (released || !Manager) {
			return;
		}
		history = createCollaborationHistory(doc as Doc, Manager, localOrigin);
		detach = target.attachCollaboration(history);
	});
	return () => {
		released = true;
		detach?.();
		history?.destroy();
		detach = null;
		history = null;
	};
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
