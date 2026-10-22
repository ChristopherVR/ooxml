/**
 * Recovery-snapshot housekeeping every binding runs: Options > Save > "Delete cached files",
 * "clear cache on close" and "cache retention" (days).
 *
 * Written once against the viewer's `storage`: with a host recovery store
 * (`ooxml-core/embedding`) the sweep lists and clears the host's copies, but only when the host
 * also provides `listRecovery`; without it nothing is swept and retention is the host's job.
 * Without a host store the library's own IndexedDB database is swept, as before.
 */
import type { DocumentStorage } from 'ooxml-core/embedding';

import { hostRecoveryStore } from './autosave-host-store';
import { deleteAutosaveSnapshot, listAutosaveSnapshots } from './autosave-store';
import type { ViewerOptions } from './options/viewer-options';
import { resolveExpiredAutosaveSnapshots } from './options/viewer-options-apply';

/** True when the active recovery store can enumerate its copies (IndexedDB, or a host with `listRecovery`). */
export function canSweepAutosaveSnapshots(storage?: DocumentStorage | null): boolean {
	const host = hostRecoveryStore(storage);
	return host ? typeof host.listRecovery === 'function' : true;
}

/** Delete every recovery snapshot of the active store; resolves with how many were deleted. */
export async function clearAutosaveSnapshots(storage?: DocumentStorage | null): Promise<number> {
	if (!canSweepAutosaveSnapshots(storage)) {
		return 0;
	}
	const snapshots = await listAutosaveSnapshots(storage);
	await Promise.all(snapshots.map((entry) => deleteAutosaveSnapshot(entry.key, storage)));
	return snapshots.length;
}

/** Delete the snapshots older than Options > Save > "cache retention"; resolves with their keys. */
export async function pruneExpiredAutosaveSnapshots(
	options: ViewerOptions,
	storage?: DocumentStorage | null,
	now: number = Date.now(),
): Promise<string[]> {
	if (!canSweepAutosaveSnapshots(storage)) {
		return [];
	}
	const expired = resolveExpiredAutosaveSnapshots(
		await listAutosaveSnapshots(storage),
		options,
		now,
	);
	await Promise.all(expired.map((key) => deleteAutosaveSnapshot(key, storage)));
	return expired;
}
