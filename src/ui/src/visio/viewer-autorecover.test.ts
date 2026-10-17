import { afterEach, describe, expect, it, vi } from 'vitest';
import type { ViewerController, ViewerState } from './controller';
import {
	ViewerAutoRecover,
	indexedDbRecoveryStore,
	memoryRecoveryStore,
	type RecoveryOffer,
} from './viewer-autorecover';

afterEach(() => vi.useRealTimers());

const bytes = (...values: number[]) => Uint8Array.from(values);

/** The slice of the controller the logic reads, with state a test can change. */
function rig(
	overrides: { dirty?: boolean; source?: boolean; loading?: boolean; busy?: boolean } = {},
) {
	const flags = { dirty: true, source: true, loading: false, busy: false, ...overrides };
	let generation = 1;
	let current: Uint8Array = bytes(1, 2, 3);
	const load = vi.fn(async (next: Uint8Array) => {
		current = next;
		generation++;
	});
	const controller = {
		get state() {
			return {
				document: {},
				loading: flags.loading,
				edit: { dirty: flags.dirty, sourceAvailable: flags.source, busy: flags.busy },
			} as unknown as ViewerState;
		},
		get sourceGeneration() {
			return generation;
		},
		get documentGeneration() {
			return generation;
		},
		exportVsdx: () => ({ bytes: current, dirty: flags.dirty }),
		load,
	} as unknown as ViewerController;
	const store = memoryRecoveryStore();
	const offers: RecoveryOffer[] = [];
	const timers = {
		callbacks: [] as (() => void)[],
		cleared: [] as unknown[],
		setInterval: vi.fn((handler: () => void) => {
			timers.callbacks.push(handler);
			return timers.callbacks.length;
		}),
		clearInterval: vi.fn((id: unknown) => void timers.cleared.push(id)),
	};
	let name = 'Plan.vsdx';
	const recover = new ViewerAutoRecover(
		controller,
		() => name,
		store,
		(offer) => offers.push(offer),
		() => 1_000,
		timers,
	);
	return {
		recover,
		store,
		offers,
		timers,
		flags,
		load,
		setBytes: (value: Uint8Array) => (current = value),
		setName: (value: string) => (name = value),
		newSource: () => generation++,
	};
}

describe('AutoRecover saving', () => {
	it('keeps a copy of a drawing with unsaved changes, under its file name', async () => {
		const { recover, store } = rig();
		recover.configure(true, 5);
		expect(await recover.saveNow()).toBe(true);
		expect(store.entries.get('Plan.vsdx')).toMatchObject({ name: 'Plan.vsdx', savedAt: 1_000 });
		expect([...store.entries.get('Plan.vsdx')!.bytes]).toEqual([1, 2, 3]);
	});

	it('names an unnamed drawing Drawing1.vsdx', async () => {
		const { recover, store, setName } = rig();
		setName('');
		recover.configure(true, 5);
		await recover.saveNow();
		expect(store.entries.has('Drawing1.vsdx')).toBe(true);
	});

	it.each([
		['the option is off', { dirty: true }, false],
		['there are no unsaved changes', { dirty: false }, true],
		['the drawing is read-only', { source: false }, true],
		['it is still opening', { loading: true }, true],
		['an edit is running', { busy: true }, true],
	])('keeps nothing when %s', async (_why, flags, enabled) => {
		const { recover, store } = rig(flags);
		recover.configure(enabled, 5);
		expect(await recover.saveNow()).toBe(false);
		expect(store.entries.size).toBe(0);
	});

	it('copies on a timer set from the option, and stops when it is off', () => {
		const { recover, timers } = rig();
		recover.configure(true, 5);
		expect(timers.setInterval).toHaveBeenCalledWith(expect.any(Function), 5 * 60_000);
		recover.configure(true, 2);
		expect(timers.cleared).toEqual([1]);
		expect(timers.setInterval).toHaveBeenLastCalledWith(expect.any(Function), 2 * 60_000);
		recover.configure(false, 2);
		expect(timers.cleared).toEqual([1, 2]);
		expect(timers.setInterval).toHaveBeenCalledTimes(2);
		expect(recover.enabled).toBe(false);
	});

	it('survives a store that fails', async () => {
		const { recover, store } = rig();
		store.put = async () => {
			throw new Error('quota');
		};
		recover.configure(true, 5);
		await expect(recover.saveNow()).resolves.toBe(false);
	});

	it('also copies when the page is hidden, and stops listening once disposed', async () => {
		const { recover, store } = rig();
		recover.configure(true, 5);
		let visibility = 'hidden';
		const doc = new EventTarget() as unknown as Document;
		Object.defineProperty(doc, 'visibilityState', { get: () => visibility });
		const dispose = recover.wire(doc);
		doc.dispatchEvent(new Event('visibilitychange'));
		await vi.waitFor(() => expect(store.entries.size).toBe(1));
		store.entries.clear();
		visibility = 'visible';
		doc.dispatchEvent(new Event('visibilitychange'));
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(store.entries.size).toBe(0);
		dispose();
		visibility = 'hidden';
		doc.dispatchEvent(new Event('visibilitychange'));
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(store.entries.size).toBe(0);
	});

	it('drops the kept copy once a copy has been saved as a file', async () => {
		const { recover, store } = rig();
		recover.configure(true, 5);
		await recover.saveNow();
		recover.saved();
		await vi.waitFor(() => expect(store.entries.size).toBe(0));
	});
});

describe('AutoRecover offering', () => {
	const kept = { name: 'Plan.vsdx', savedAt: 5_000, bytes: bytes(9, 9) };

	it('ignores a storage lookup that finishes after another drawing opens', async () => {
		const { recover, store, offers, newSource } = rig({ dirty: false });
		let finish!: (entry: typeof kept) => void;
		store.get = () =>
			new Promise((resolve) => {
				finish = resolve;
			});
		recover.opened();
		newSource();
		finish(kept);
		await Promise.resolve();
		expect(offers).toHaveLength(0);
	});

	it('refuses to restore an offer after the user edits the drawing', async () => {
		const { recover, store, offers, flags, load } = rig({ dirty: false });
		await store.put(kept);
		recover.opened();
		await vi.waitFor(() => expect(offers).toHaveLength(1));
		flags.dirty = true;
		await expect(offers[0]!.restore()).rejects.toThrow('The drawing changed');
		expect(load).not.toHaveBeenCalled();
		expect(store.entries.has(kept.name)).toBe(true);
	});

	it('offers a kept copy when the same file opens, even with the option off', async () => {
		const { recover, store, offers } = rig({ dirty: false });
		await store.put(kept);
		recover.opened();
		await vi.waitFor(() => expect(offers).toHaveLength(1));
		expect(offers[0]).toMatchObject({ name: 'Plan.vsdx', savedAt: 5_000 });
	});

	it('asks again each time the file is opened, since Keep for later left the copy', async () => {
		const { recover, store, offers } = rig({ dirty: false });
		await store.put(kept);
		recover.opened();
		await vi.waitFor(() => expect(offers).toHaveLength(1));
		recover.opened();
		await vi.waitFor(() => expect(offers).toHaveLength(2));
	});

	it('does not look when the drawing is unsaved, read-only or still opening', async () => {
		for (const flags of [{ dirty: true }, { source: false }, { loading: true }]) {
			const { recover, store, offers } = rig({ dirty: false, ...flags });
			await store.put(kept);
			recover.opened();
			await new Promise((resolve) => setTimeout(resolve, 20));
			expect(offers).toHaveLength(0);
		}
	});

	it('leaves a copy that is the same as the opened file, offered to nobody and kept', async () => {
		const { recover, store, offers } = rig({ dirty: false });
		await store.put({ ...kept, bytes: bytes(1, 2, 3) });
		recover.opened();
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(offers).toHaveLength(0);
		expect(store.entries.has('Plan.vsdx')).toBe(true);
	});

	it('only finds a copy of the same file name', async () => {
		const { recover, store, offers, setName } = rig({ dirty: false });
		await store.put(kept);
		setName('Other.vsdx');
		recover.opened();
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(offers).toHaveLength(0);
	});

	it('restores by opening the kept bytes, and discards by deleting them', async () => {
		const { recover, store, offers, load } = rig({ dirty: false });
		await store.put(kept);
		recover.opened();
		await vi.waitFor(() => expect(offers).toHaveLength(1));
		await offers[0]!.restore();
		expect([...load.mock.calls[0]![0]]).toEqual([9, 9]);
		expect(store.entries.has('Plan.vsdx')).toBe(true);
		await offers[0]!.discard();
		expect(store.entries.has('Plan.vsdx')).toBe(false);
	});

	it('does not offer the restored copy back to itself', async () => {
		const { recover, store, offers, setBytes, newSource } = rig({ dirty: false });
		await store.put(kept);
		recover.opened();
		await vi.waitFor(() => expect(offers).toHaveLength(1));
		await offers[0]!.restore();
		setBytes(kept.bytes);
		newSource();
		recover.opened();
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(offers).toHaveLength(1);
		expect(store.entries.has('Plan.vsdx')).toBe(true);
	});

	it('survives a store that cannot be read', async () => {
		const { recover, store, offers } = rig({ dirty: false });
		store.get = async () => {
			throw new Error('blocked');
		};
		recover.opened();
		await new Promise((resolve) => setTimeout(resolve, 20));
		expect(offers).toHaveLength(0);
	});
});

describe('stores', () => {
	it('has no IndexedDB store where the browser has none', () => {
		expect(indexedDbRecoveryStore(undefined)).toBeUndefined();
	});

	it('the memory store keeps its own copy of the bytes', async () => {
		const store = memoryRecoveryStore();
		const original = bytes(1, 2);
		await store.put({ name: 'a', savedAt: 1, bytes: original });
		original[0] = 99;
		expect([...(await store.get('a'))!.bytes]).toEqual([1, 2]);
		await store.delete('a');
		expect(await store.get('a')).toBeUndefined();
	});
});
