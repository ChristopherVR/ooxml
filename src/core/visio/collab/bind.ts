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
import { applyFixups } from './fixups';
import { DrawingKeys } from './keys';
import { readSharedDrawing } from './read';
import { SchemaVersionError, isSharedDrawingEmpty, sharedDrawing, sharedTypes } from './schema';
import { snapshotDrawing } from './snapshot';
import { writeSharedDrawing } from './write';

/** Origin of the binding's per-edit transactions (the ones its undo manager tracks). */
export const VISIO_EDIT_ORIGIN: unique symbol = Symbol('ooxml-core:visio:edit');
/** Origin of seeding writes, which are not undoable. */
export const VISIO_SYNC_ORIGIN: unique symbol = Symbol('ooxml-core:visio:sync');
/** Origin of the fix-ups that persist renumbered ids: not a local edit, so never in an undo history. */
export const VISIO_FIXUP_ORIGIN: unique symbol = Symbol('ooxml-core:visio:fixup');

/** What differs between the drawing last applied and the next one, when only page content does. */
export interface DrawingChange {
	/** Ids (`Page/@ID` in the package handed over) of the pages whose content changed. */
	pages: string[];
}

/** What the binding needs from the editor that shows the drawing. DOM-free. */
export interface DrawingEditor {
	/** The package of the open drawing, or undefined when none can be read now. */
	exportBytes(): Uint8Array | undefined;
	/**
	 * Replaces the open drawing with `bytes` as a step outside the local undo history. `change`
	 * is given only when the open drawing is exactly the one last applied and just the content of
	 * these pages differs (same page list, masters, styles, theme, media and other parts): the
	 * editor may then read only those pages. Without it, read the whole package.
	 */
	applyRemote(bytes: Uint8Array, change?: DrawingChange): Promise<void>;
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
	/**
	 * Runs `run` once after the current burst of remote transactions (default: a zero-delay timer).
	 * Remote changes that arrive before it runs share one rebuild. `idle()` waits for it, so a test
	 * that injects a manual scheduler must call `run` before awaiting `idle()`.
	 */
	schedule?: (run: () => void) => void;
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
	let incompatible = false;
	let queued = false;
	let queuedDone: Promise<void> | undefined;
	const schedule =
		options.schedule ??
		((run: () => void) => {
			setTimeout(run, 0);
		});
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
		// Persist renumbered ids (what every peer computes alike); the drawing itself is unchanged.
		if (resolved.fixups.length && host.canWrite())
			doc.transact(() => applyFixups(shared, resolved.fixups), VISIO_FIXUP_ORIGIN);
		const version = editor.version();
		// Nothing the room holds changed the drawing this editor already shows.
		if (keys.applied?.signature === resolved.signature && keys.applied.version === version) return;
		// Only page content differs from what the editor shows: let it read just those pages.
		const before = keys.applied;
		const change: DrawingChange | undefined =
			before && before.version === version && before.base === resolved.baseSignature
				? {
						pages: resolved.pages
							.filter((page) => before.pages.get(page.key) !== page.contentSignature)
							.map((page) => page.id),
					}
				: undefined;
		const composed = await composeDrawing(resolved, editor.exportBytes(), keys);
		if (disposed) return;
		if (editor.busy() || editor.version() !== version) {
			remoteDirty = true; // The person edited meanwhile: publish that, then rebuild again.
			return;
		}
		try {
			await editor.applyRemote(composed.bytes, change);
		} catch (error) {
			if (editor.busy()) remoteDirty = true;
			else report(error);
			return;
		}
		composed.commit();
		keys.applied = {
			signature: composed.signature,
			version: editor.version(),
			base: resolved.baseSignature,
			pages: new Map(resolved.pages.map((page) => [page.key, page.contentSignature])),
		};
	};

	const blocked = (): boolean => disposed || incompatible || !settled || editor.busy();
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
				// A room with another schema version is left alone: no writes, no rebuilds.
				if (error instanceof SchemaVersionError) incompatible = true;
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

	const queueRemote = (): void => {
		if (queued) return;
		queued = true;
		queuedDone = new Promise<void>((resolve) =>
			schedule(() => {
				queued = false;
				remoteDirty = true;
				resolve();
				pump();
			}),
		);
	};

	const afterTransaction = (transaction: Y.Transaction): void => {
		if (disposed) return;
		const origin = transaction.origin;
		if (
			origin === VISIO_EDIT_ORIGIN ||
			origin === VISIO_SYNC_ORIGIN ||
			origin === VISIO_FIXUP_ORIGIN
		)
			return;
		const changed = transaction.changed as Map<unknown, unknown>;
		if (!types.some((type) => changed.has(type))) return;
		if (!settled) {
			if (!isSharedDrawingEmpty(shared)) settle();
			return;
		}
		queueRemote();
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
		while (running || queued) {
			if (queued) await queuedDone;
			if (running) await running;
		}
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
