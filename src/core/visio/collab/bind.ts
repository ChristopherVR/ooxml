// Binds an open Visio drawing to a shared drawing. A local edit becomes one Yjs transaction tagged
// VISIO_EDIT_ORIGIN that writes only what the person changed; a transaction from anyone else (or
// from this binding's undo manager) rebuilds the package from the room and hands it to the editor
// outside its history. Undo is a Y.UndoManager tracking only local edits, so it never reverts a
// collaborator's work: route Ctrl+Z through the binding.
//
// The editor works on whole packages and its edits are asynchronous, so the binding is a small
// work loop: publish the local edit first (a rebuild would drop it), then rebuild from the room.
import * as Y from 'yjs';
import type { CollabSession } from '../../collab/session';
import { composeDrawing } from './compose';
import { DrawingKeys } from './keys';
import { readSharedDrawing } from './read';
import { isSharedDrawingEmpty, sharedDrawing, sharedTypes } from './schema';
import { snapshotDrawing } from './snapshot';
import { writeSharedDrawing } from './write';

/** Origin of the binding's per-edit transactions (the ones its undo manager tracks). */
export const VISIO_EDIT_ORIGIN: unique symbol = Symbol('ooxml-core:visio:edit');
/** Origin of seeding writes, which are not undoable. */
export const VISIO_SYNC_ORIGIN: unique symbol = Symbol('ooxml-core:visio:sync');

/** What the binding needs from the editor that shows the drawing. DOM-free. */
export interface DrawingEditor {
	/** The package of the open drawing, or undefined when none can be read now. */
	exportBytes(): Uint8Array | undefined;
	/** Replaces the open drawing with `bytes` as a step outside the local undo history. */
	applyRemote(bytes: Uint8Array): Promise<void>;
	/** Called after the person edited or loaded a drawing (never for `applyRemote`). */
	onLocalChange(listener: () => void): () => void;
	/** Called whenever the editor may have stopped being busy. */
	onIdle(listener: () => void): () => void;
	/** True while the editor cannot export or accept a drawing. */
	busy(): boolean;
	/** Changes whenever the open drawing does. */
	version(): number;
}

/** What the binding needs from a collab session. */
export type DrawingCollabHost = Pick<CollabSession<object>, 'doc' | 'synced' | 'canWrite' | 'on'>;

export interface DrawingBindingOptions {
	/** Seed an empty room with the open drawing (default true; false for `intent: 'join'`). */
	seed?: boolean;
	/** A change could not be shared or applied. */
	onError?: (error: unknown) => void;
	/** Part of every key this peer mints; defaults to the Yjs client id. */
	keyTag?: string;
}

export interface DrawingBinding {
	/** Undoes the last local edit (never a collaborator's); resolves false when there is none. */
	undo: () => Promise<boolean>;
	redo: () => Promise<boolean>;
	canUndo: () => boolean;
	canRedo: () => boolean;
	/** Called whenever the undo or redo stack changes. */
	onHistoryChange: (listener: () => void) => () => void;
	/** Resolves when all queued publishing and applying has finished or is waiting on the editor. */
	idle: () => Promise<void>;
	dispose: () => void;
}

export function bindDrawingSession(
	host: DrawingCollabHost,
	editor: DrawingEditor,
	options: DrawingBindingOptions = {},
): DrawingBinding {
	const { doc } = host;
	const shared = sharedDrawing(doc);
	const types = sharedTypes(shared);
	const keys = new DrawingKeys(options.keyTag ?? doc.clientID.toString(36));
	const undoManager = new Y.UndoManager(types, {
		trackedOrigins: new Set([VISIO_EDIT_ORIGIN]),
		captureTimeout: 0,
	});
	const historyListeners = new Set<() => void>();
	const historyChanged = (): void => {
		for (const listener of [...historyListeners]) listener();
	};
	for (const event of ['stack-item-added', 'stack-item-popped', 'stack-cleared'] as const)
		undoManager.on(event, historyChanged);

	let disposed = false;
	let settled = false;
	let localDirty = false;
	let remoteDirty = false;
	let running: Promise<void> | undefined;
	const report = (error: unknown): void => options.onError?.(error);

	const publish = async (): Promise<void> => {
		const bytes = editor.exportBytes();
		if (!bytes) return;
		keys.seeding = keys.known.size === 0 && isSharedDrawingEmpty(shared);
		const snapshot = await snapshotDrawing(bytes, keys);
		if (disposed) return;
		// Only a drawing that fills an empty room is a seed; anything later is an undoable edit.
		const seed = isSharedDrawingEmpty(shared);
		doc.transact(
			() => writeSharedDrawing(shared, snapshot, keys),
			seed ? VISIO_SYNC_ORIGIN : VISIO_EDIT_ORIGIN,
		);
		undoManager.stopCapturing();
	};

	const rebuild = async (): Promise<void> => {
		const resolved = readSharedDrawing(shared);
		if (!resolved) return;
		const version = editor.version();
		const composed = await composeDrawing(resolved, editor.exportBytes(), keys);
		if (disposed) return;
		if (editor.busy() || editor.version() !== version) {
			remoteDirty = true; // The person edited meanwhile: publish that, then rebuild again.
			return;
		}
		try {
			await editor.applyRemote(composed.bytes);
		} catch (error) {
			if (editor.busy()) remoteDirty = true;
			else report(error);
			return;
		}
		composed.commit();
	};

	const blocked = (): boolean => disposed || !settled || editor.busy();
	const pending = (): boolean =>
		!blocked() && ((localDirty && host.canWrite()) || (remoteDirty && !localDirty));

	const loop = async (): Promise<void> => {
		while (pending()) {
			try {
				if (localDirty) {
					localDirty = false;
					await publish();
				} else if (remoteDirty) {
					remoteDirty = false;
					await rebuild();
				}
			} catch (error) {
				report(error);
			}
		}
	};
	const pump = (): void => {
		if (running || !pending()) return;
		running = loop().finally(() => {
			running = undefined;
			pump();
		});
	};

	const settle = (): void => {
		if (disposed || settled) return;
		if (isSharedDrawingEmpty(shared)) {
			if (!host.canWrite()) return;
			settled = true;
			if (options.seed !== false) {
				localDirty = editor.exportBytes() !== undefined || editor.busy();
			}
		} else {
			settled = true;
			localDirty = false; // The room wins over a drawing opened before joining.
			remoteDirty = true;
		}
		pump();
	};

	const afterTransaction = (transaction: Y.Transaction): void => {
		if (disposed) return;
		const origin = transaction.origin;
		if (origin === VISIO_EDIT_ORIGIN || origin === VISIO_SYNC_ORIGIN) return;
		const changed = transaction.changed as Map<unknown, unknown>;
		if (!types.some((type) => changed.has(type))) return;
		if (!settled) {
			if (!isSharedDrawingEmpty(shared)) settle();
			return;
		}
		remoteDirty = true;
		pump();
	};
	doc.on('afterTransaction', afterTransaction);

	const stops = [
		editor.onLocalChange(() => {
			if (disposed) return;
			localDirty = true;
			pump();
		}),
		editor.onIdle(pump),
		host.on('ready', settle),
		host.on('synced', () => {
			settle();
			pump();
		}),
	];
	if (host.synced || host.canWrite()) settle();

	const idle = async (): Promise<void> => {
		while (running) await running;
	};
	const replay = async (kind: 'undo' | 'redo'): Promise<boolean> => {
		await idle();
		const stack = kind === 'undo' ? undoManager.undoStack : undoManager.redoStack;
		if (disposed || !stack.length) return false;
		if (kind === 'undo') undoManager.undo();
		else undoManager.redo();
		await idle();
		return true;
	};

	return {
		undo: () => replay('undo'),
		redo: () => replay('redo'),
		canUndo: () => undoManager.undoStack.length > 0,
		canRedo: () => undoManager.redoStack.length > 0,
		onHistoryChange: (listener) => {
			historyListeners.add(listener);
			return () => historyListeners.delete(listener);
		},
		idle,
		dispose: () => {
			if (disposed) return;
			disposed = true;
			doc.off('afterTransaction', afterTransaction);
			for (const stop of stops) stop();
			undoManager.destroy();
			historyListeners.clear();
		},
	};
}
