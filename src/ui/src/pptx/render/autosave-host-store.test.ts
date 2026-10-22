// @vitest-environment jsdom
import type { DocumentRecoveryEntry, DocumentStorage } from 'ooxml-core/embedding';
import { describe, expect, it, vi } from 'vitest';

import { discardAutosaveRecovery, probeAutosaveRecovery } from './autosave-recovery';
import {
	deleteAutosaveSnapshot,
	getAutosaveSnapshot,
	listAutosaveSnapshots,
	saveAutosaveSnapshot,
	wasWrittenByThisPage,
} from './autosave-store';

function hostStorage(withList = true) {
	const entries = new Map<string, DocumentRecoveryEntry>();
	const storage: DocumentStorage = {
		save: vi.fn(async () => undefined),
		readRecovery: vi.fn(async (id: string) => entries.get(id)),
		writeRecovery: vi.fn(async (entry: DocumentRecoveryEntry) => void entries.set(entry.id, entry)),
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

describe('autosave store with a host recovery store', () => {
	it('writes, reads, lists and clears through the host instead of IndexedDB', async () => {
		const { storage, entries } = hostStorage();
		const bytes = new Uint8Array([1, 2, 3]);
		expect(await saveAutosaveSnapshot('decks/a.pptx', bytes, storage)).toBe(true);
		expect(storage.writeRecovery).toHaveBeenCalledTimes(1);
		const stored = entries.get('decks/a.pptx')!;
		expect(stored.bytes).toBe(bytes);
		expect(wasWrittenByThisPage('decks/a.pptx', stored.savedAt)).toBe(true);

		const record = await getAutosaveSnapshot('decks/a.pptx', storage);
		expect(record).toEqual({
			key: 'decks/a.pptx',
			data: bytes,
			timestamp: stored.savedAt,
			size: 3,
		});
		expect(await listAutosaveSnapshots(storage)).toEqual([
			{ key: 'decks/a.pptx', timestamp: stored.savedAt, size: 3 },
		]);
		expect(await deleteAutosaveSnapshot('decks/a.pptx', storage)).toBe(true);
		expect(storage.clearRecovery).toHaveBeenCalledWith('decks/a.pptx');
		expect(await getAutosaveSnapshot('decks/a.pptx', storage)).toBeUndefined();
	});

	it('lists nothing when the host does not list its copies', async () => {
		const { storage } = hostStorage(false);
		await saveAutosaveSnapshot('a.pptx', new Uint8Array([1]), storage);
		expect(await listAutosaveSnapshots(storage)).toEqual([]);
	});

	it('propagates a failing host write', async () => {
		const { storage } = hostStorage();
		storage.writeRecovery = () => Promise.reject(new Error('quota'));
		await expect(saveAutosaveSnapshot('a.pptx', new Uint8Array([1]), storage)).rejects.toThrow(
			'quota',
		);
	});

	it('offers and discards a host snapshot even without IndexedDB', async () => {
		const { storage, entries } = hostStorage();
		const savedAt = Date.now() - 60_000;
		entries.set('b.pptx', { id: 'b.pptx', savedAt, bytes: new Uint8Array([9]) });
		const offer = await probeAutosaveRecovery('b.pptx', Date.now(), 'B', storage);
		expect(offer?.record).toMatchObject({ key: 'b.pptx', timestamp: savedAt, size: 1 });
		await discardAutosaveRecovery(offer!.record, storage);
		expect(entries.has('b.pptx')).toBe(false);
	});

	it('treats a host store missing a recovery method as none (IndexedDB stays the default)', async () => {
		const storage: DocumentStorage = { save: async () => undefined, readRecovery: vi.fn() };
		// jsdom has no IndexedDB: the probe reports nothing instead of reading the host.
		expect(await probeAutosaveRecovery('c.pptx', Date.now(), undefined, storage)).toBeNull();
		expect(storage.readRecovery).not.toHaveBeenCalled();
	});
});
