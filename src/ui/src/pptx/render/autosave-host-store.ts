/**
 * Autosave recovery copies kept by the host instead of the library's IndexedDB database.
 *
 * When the viewer's `storage` prop provides the recovery methods of `ooxml-core/embedding`
 * (`readRecovery`, `writeRecovery`, `clearRecovery`, optionally `listRecovery`), the autosave
 * store functions of `autosave-store.ts` route here. The records keep the store's own
 * {@link AutosaveRecord} shape, so every caller (recovery prompt, version history, session
 * restore) reads them the same way. Keys are the host-supplied `filePath`, as before.
 */
import {
	recoveryStoreOf,
	type DocumentRecoveryStore,
	type DocumentStorage,
} from 'ooxml-core/embedding';

/** A stored recovery snapshot (the IndexedDB record shape). */
export interface AutosaveRecord {
	key: string;
	data: Uint8Array;
	timestamp: number;
	size: number;
}

/** The host's recovery store, when `storage` has one; undefined keeps IndexedDB. */
export const hostRecoveryStore = (
	storage: DocumentStorage | null | undefined,
): DocumentRecoveryStore | undefined => recoveryStoreOf(storage);

export async function writeHostSnapshot(
	store: DocumentRecoveryStore,
	key: string,
	data: Uint8Array,
	timestamp: number,
): Promise<void> {
	await store.writeRecovery({ id: key, savedAt: timestamp, bytes: data });
}

export async function readHostSnapshot(
	store: DocumentRecoveryStore,
	key: string,
): Promise<AutosaveRecord | undefined> {
	const entry = await store.readRecovery(key);
	if (!entry) {
		return undefined;
	}
	return {
		key: entry.id,
		data: entry.bytes,
		timestamp: entry.savedAt,
		size: entry.bytes.byteLength,
	};
}

/** Every kept copy, without bytes; empty when the host does not list its copies. */
export async function listHostSnapshots(
	store: DocumentRecoveryStore,
): Promise<Array<{ key: string; timestamp: number; size: number }>> {
	const entries = (await store.listRecovery?.()) ?? [];
	return entries.map((entry) => ({ key: entry.id, timestamp: entry.savedAt, size: entry.size }));
}
