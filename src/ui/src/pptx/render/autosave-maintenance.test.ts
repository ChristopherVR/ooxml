// @vitest-environment jsdom
import type { DocumentRecoveryEntry, DocumentStorage } from 'ooxml-core/embedding';
import { describe, expect, it, vi } from 'vitest';

import { getLocalStorageUsageSummary, clearAllLocalViewerData } from './account';
import {
	canSweepAutosaveSnapshots,
	clearAutosaveSnapshots,
	pruneExpiredAutosaveSnapshots,
} from './autosave-maintenance';
import { listBackstageRecentFiles, readBackstageRecentFile } from './backstage';
import { DEFAULT_VIEWER_OPTIONS } from './options/viewer-options';

const DAY = 24 * 60 * 60 * 1000;

function hostStorage(withList = true) {
	const entries = new Map<string, DocumentRecoveryEntry>();
	const storage: DocumentStorage = {
		save: async () => undefined,
		readRecovery: vi.fn(async (id: string) => entries.get(id)),
		writeRecovery: async (entry) => void entries.set(entry.id, entry),
		clearRecovery: vi.fn(async (id: string) => void entries.delete(id)),
		...(withList
			? {
					listRecovery: async () =>
						[...entries.values()].map((entry) => ({
							id: entry.id,
							savedAt: entry.savedAt,
							size: entry.bytes.byteLength,
						})),
				}
			: {}),
	};
	return { storage, entries };
}

const put = (entries: Map<string, DocumentRecoveryEntry>, id: string, savedAt: number) =>
	entries.set(id, { id, savedAt, bytes: new Uint8Array([1, 2]) });

describe('recovery housekeeping with a host recovery store (jsdom has no IndexedDB)', () => {
	it('clears every host copy when the host lists them', async () => {
		const { storage, entries } = hostStorage();
		put(entries, 'a.pptx', 1);
		put(entries, 'b.pptx', 2);
		expect(canSweepAutosaveSnapshots(storage)).toBe(true);
		expect(await clearAutosaveSnapshots(storage)).toBe(2);
		expect(entries.size).toBe(0);
	});

	it('prunes only copies older than the retention option', async () => {
		const { storage, entries } = hostStorage();
		const now = 100 * DAY;
		put(entries, 'old.pptx', now - 40 * DAY);
		put(entries, 'new.pptx', now - DAY);
		const options = {
			...DEFAULT_VIEWER_OPTIONS,
			save: { ...DEFAULT_VIEWER_OPTIONS.save, cacheRetentionDays: 30 },
		};
		expect(await pruneExpiredAutosaveSnapshots(options, storage, now)).toEqual(['old.pptx']);
		expect([...entries.keys()]).toEqual(['new.pptx']);
	});

	it('leaves retention to a host that cannot list its copies', async () => {
		const { storage, entries } = hostStorage(false);
		put(entries, 'a.pptx', 1);
		expect(canSweepAutosaveSnapshots(storage)).toBe(false);
		expect(await clearAutosaveSnapshots(storage)).toBe(0);
		expect(await pruneExpiredAutosaveSnapshots(DEFAULT_VIEWER_OPTIONS, storage)).toEqual([]);
		expect(storage.clearRecovery).not.toHaveBeenCalled();
	});

	it('lists and reads Recent files and Account usage from the host', async () => {
		const { storage, entries } = hostStorage();
		put(entries, 'decks/q3.pptx', 5);
		const recent = await listBackstageRecentFiles(undefined, undefined, storage);
		expect(recent).toMatchObject([{ key: 'decks/q3.pptx', name: 'q3.pptx', size: 2 }]);
		expect(await readBackstageRecentFile('decks/q3.pptx', storage)).toEqual(new Uint8Array([1, 2]));
		expect(await getLocalStorageUsageSummary(storage)).toEqual({
			presentationCount: 1,
			totalBytes: 2,
		});
		await clearAllLocalViewerData(storage);
		expect(entries.size).toBe(0);
	});

	it('keeps the no-IndexedDB answer without a host store', async () => {
		expect(await listBackstageRecentFiles()).toEqual([]);
		expect(await readBackstageRecentFile('x')).toBeUndefined();
	});
});
