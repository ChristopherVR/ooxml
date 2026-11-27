// @vitest-environment node
import { IDBFactory } from 'fake-indexeddb';
import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { createAddInPreferencesStore } from './add-in-preferences-store';

const name = 'add-ins-test';
const limit = 16 * 1024 * 1024;

beforeEach(() => vi.stubGlobal('indexedDB', new IDBFactory()));
afterEach(() => vi.unstubAllGlobals());

it('enforces one profile budget across concurrent writes and allows replacement and removal', async () => {
	const store = createAddInPreferencesStore(name);
	const anotherConnection = createAddInPreferencesStore(name);
	const file = new Uint8Array(limit / 2 + 1);
	const results = await Promise.allSettled([
		store.write('word', 'one', file),
		anotherConnection.write('excel', 'two', file),
	]);
	const word = await store.read('word');
	const excel = await store.read('excel');
	expect(results.map(({ status }) => status).sort()).toEqual(['fulfilled', 'rejected']);
	expect(word.length + excel.length).toBe(1);
	const [host, id, other] = word.length
		? (['word', 'one', 'excel'] as const)
		: (['excel', 'two', 'word'] as const);
	await store.write(host, id, new Uint8Array(limit));
	await expect(store.write(other, 'extra', Uint8Array.of(1))).rejects.toThrow(/limit/i);
	await store.write(host, id, undefined);
	await store.write(other, 'extra', Uint8Array.of(1));
	expect(await store.read(other)).toEqual([Uint8Array.of(1)]);
});

it('reads legacy records within the same budget across products', async () => {
	await new Promise<void>((resolve, reject) => {
		const request = indexedDB.open(name, 1);
		request.onupgradeneeded = () => {
			const store = request.result.createObjectStore('installs');
			store.put(new Uint8Array(limit - 1), 'excel:large');
			store.put(Uint8Array.of(1, 2), 'word:oversize');
			store.put(Uint8Array.of(3), 'word:small');
		};
		request.onsuccess = () => {
			request.result.close();
			resolve();
		};
		request.onerror = () => reject(request.error);
	});
	const store = createAddInPreferencesStore(name);
	expect((await store.read('excel')).map(({ byteLength }) => byteLength)).toEqual([limit - 1]);
	expect(await store.read('word')).toEqual([Uint8Array.of(3)]);
});
