// Binds an xlsx edit session to a shared workbook. Every local edit becomes one Yjs transaction
// tagged XLSX_EDIT_ORIGIN that writes only what differs; remote transactions are applied to the
// session through `applyExternal` (outside its history). Undo is a Y.UndoManager tracking only
// local edits, so it never reverts a collaborator's work: route Ctrl+Z through the binding.
import * as Y from 'yjs';
import type { CollabSession } from '../../collab/session';
import type { CellRange } from '../address';
import type { EditSession, WorkbookChange } from '../edit/types';
import { applyShared } from './apply';
import { collectDirty, emptyDirty, isDirty } from './dirty';
import { type XlsxPresence, type RemoteSelection, resolveSelections } from './presence';
import { formatRange } from '../address';
import { sharedTypes, sharedWorkbook } from './schema';
import { SheetKeys, sheetEntries } from './sheets';
import { type WriteScope, writeWorkbook } from './write';

/** Origin of the binding's per-edit transactions (the ones its undo manager tracks). */
export const XLSX_EDIT_ORIGIN: unique symbol = Symbol('ooxml-core:xlsx:edit');
/** Origin of seeding and catch-up writes, which are not undoable. */
export const XLSX_SYNC_ORIGIN: unique symbol = Symbol('ooxml-core:xlsx:sync');

/** What the binding needs from a collab session (a `CollabSession<XlsxPresence>` satisfies it). */
export type WorkbookCollabHost = Pick<
	CollabSession<XlsxPresence>,
	'doc' | 'synced' | 'canWrite' | 'on' | 'updatePresence' | 'peers'
>;

export interface WorkbookBindingOptions {
	/** A remote change could not be applied (it is skipped; the room is still converged). */
	onError?: (error: unknown) => void;
}

export interface WorkbookBinding {
	/** Undo the last local edit (never a collaborator's). Returns false when there is none. */
	undo: () => boolean;
	redo: () => boolean;
	canUndo: () => boolean;
	canRedo: () => boolean;
	/** Publish the local selection (sheet index and range) to collaborators. */
	setSelection: (sheet: number, range: CellRange) => void;
	/** Collaborators' selections, resolved to local sheet indexes. */
	remoteSelections: () => RemoteSelection[];
	/** The shared key of a local sheet, and back. */
	sheetKey: (sheet: number) => string | undefined;
	sheetIndex: (key: string) => number | undefined;
	dispose: () => void;
}

const WHOLE_WORKBOOK = new Set(['sheets', 'structure', 'names', 'batch', 'undo', 'redo']);

/** The part of the shared workbook a local change has to reconcile. */
export function scopeOf(change: WorkbookChange): WriteScope {
	if (change.structural || change.sheet === undefined || WHOLE_WORKBOOK.has(change.kind))
		return { kind: 'all' };
	if ((change.kind === 'cells' || change.kind === 'format') && change.ranges?.length)
		return { kind: 'sheet', sheet: change.sheet, ranges: change.ranges };
	return { kind: 'sheet', sheet: change.sheet };
}

export function bindWorkbookSession(
	host: WorkbookCollabHost,
	edit: EditSession,
	options: WorkbookBindingOptions = {},
): WorkbookBinding {
	const { doc } = host;
	const shared = sharedWorkbook(doc);
	const types = sharedTypes(shared);
	const keys = new SheetKeys(doc.clientID);
	const workbook = edit.workbook;
	const undoManager = new Y.UndoManager(types, {
		trackedOrigins: new Set([XLSX_EDIT_ORIGIN]),
		captureTimeout: 0,
	});
	let disposed = false;
	let settled = false;
	/** Local edits made while writes were blocked; published in full once they are allowed. */
	let unsent = false;
	let replaying: 'undo' | 'redo' | undefined;
	let dirty = emptyDirty();
	/** The room was empty but an earlier writer is present: wait for its workbook, then adopt. */
	let waiting = false;
	const joined = Date.now();
	host.updatePresence({ joined });
	/** A writer that joined before this binding (ties broken by client id) and will seed the room. */
	const earlierWriter = (): boolean =>
		host
			.peers()
			.some(
				(peer) =>
					peer.joined !== undefined &&
					peer.role !== 'viewer' &&
					(peer.joined < joined || (peer.joined === joined && peer.clientId < doc.clientID)),
			);

	const write = (scope: WriteScope, origin: symbol): void => {
		doc.transact(() => writeWorkbook(shared, keys, workbook, scope), origin);
		undoManager.stopCapturing();
	};
	const apply = (all: boolean, kind: 'remote' | 'undo' | 'redo' = 'remote'): void => {
		const pending = all ? emptyDirty(true) : dirty;
		dirty = emptyDirty();
		try {
			edit.applyExternal(() => {
				const change = applyShared(shared, keys, workbook, pending);
				if (change) change.kind = kind;
				return change;
			});
		} catch (error) {
			options.onError?.(error);
		}
	};

	const settle = (): void => {
		if (disposed) return;
		if (shared.sheets.size === 0) {
			if (!host.canWrite()) return;
			// Synced to an empty room while an earlier writer has not seeded yet (its gate is still
			// closed): seeding now would make that writer adopt this workbook instead of its own.
			waiting = !settled && earlierWriter();
			if (waiting) return;
			write({ kind: 'all' }, XLSX_SYNC_ORIGIN);
		} else if (settled && unsent && host.canWrite()) write({ kind: 'all' }, XLSX_SYNC_ORIGIN);
		else if (!settled) apply(true);
		settled = true;
		waiting = false;
		unsent = false;
	};

	const observers = types.map((type, i) => {
		const top = (['meta', 'styles', 'names', 'sheets'] as const)[i] ?? 'sheets';
		const observer = (
			events: Y.YEvent<Y.AbstractType<unknown>>[],
			transaction: Y.Transaction,
		): void => {
			const origin = transaction.origin;
			if (origin === XLSX_EDIT_ORIGIN || origin === XLSX_SYNC_ORIGIN) return;
			collectDirty(dirty, top, events);
		};
		type.observeDeep(observer);
		return () => type.unobserveDeep(observer);
	});
	const afterTransaction = (transaction: Y.Transaction): void => {
		if (disposed || !isDirty(dirty)) return;
		if (!settled) {
			dirty = emptyDirty(); // Adoption reads the whole room.
			if (waiting && shared.sheets.size > 0) settle();
			return;
		}
		const kind = transaction.origin === undoManager ? (replaying ?? 'undo') : 'remote';
		apply(false, kind);
	};
	doc.on('afterTransaction', afterTransaction);

	const stopEdits = edit.onChange((change) => {
		if (disposed || change.external || !settled) return;
		if (!host.canWrite()) {
			unsent = true;
			return;
		}
		write(scopeOf(change), XLSX_EDIT_ORIGIN);
	});
	const stopReady = host.on('ready', settle);
	const stopSynced = host.on('synced', (synced) => {
		if (synced && !host.canWrite()) settle();
	});
	// The earlier writer left before seeding: this binding seeds after all.
	const stopPeers = host.on('peers', () => {
		if (waiting && !settled) settle();
	});
	if (host.synced || host.canWrite()) settle();

	const replay = (kind: 'undo' | 'redo'): boolean => {
		const stack = kind === 'undo' ? undoManager.undoStack : undoManager.redoStack;
		if (disposed || !stack.length) return false;
		replaying = kind;
		try {
			if (kind === 'undo') undoManager.undo();
			else undoManager.redo();
		} finally {
			replaying = undefined;
		}
		return true;
	};
	const sheetKey = (index: number): string | undefined => {
		const sheet = workbook.sheets[index];
		return sheet && keys.find(sheet, sheetEntries(shared.sheets));
	};
	const sheetIndex = (key: string): number | undefined => {
		const entries = sheetEntries(shared.sheets);
		const index = workbook.sheets.findIndex((sheet) => keys.find(sheet, entries) === key);
		return index >= 0 ? index : undefined;
	};

	return {
		undo: () => replay('undo'),
		redo: () => replay('redo'),
		canUndo: () => undoManager.undoStack.length > 0,
		canRedo: () => undoManager.redoStack.length > 0,
		setSelection: (sheet, range) => {
			const key = sheetKey(sheet);
			if (key !== undefined) host.updatePresence({ sheet: key, range: formatRange(range) });
		},
		remoteSelections: () => resolveSelections(host.peers(), sheetIndex),
		sheetKey,
		sheetIndex,
		dispose: () => {
			if (disposed) return;
			disposed = true;
			for (const stop of observers) stop();
			doc.off('afterTransaction', afterTransaction);
			stopEdits();
			stopReady();
			stopSynced();
			stopPeers();
			undoManager.destroy();
		},
	};
}
