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

const seed = (records: Record<string, Uint8Array>) =>
	new Promise<void>((resolve, reject) => {
		const request = indexedDB.open(name, 1);
		request.onupgradeneeded = () => {
			const store = request.result.createObjectStore('installs');
			for (const [key, bytes] of Object.entries(records)) store.put(bytes, key);
		};
		request.onsuccess = () => {
			request.result.close();
			resolve();
		};
		request.onerror = () => reject(request.error);
	});

it('reads legacy records within each product budget, so one product never hides another', async () => {
	await seed({
		'excel:large': new Uint8Array(limit / 2),
		'word:large': new Uint8Array(limit - 1),
		'word:small': Uint8Array.of(3),
		...Object.fromEntries(
			Array.from({ length: 40 }, (_, index) => [`excel:many-${index}`, Uint8Array.of(1)]),
		),
	});
	const store = createAddInPreferencesStore(name);
	const excel = await store.read('excel');
	expect(excel).toHaveLength(32);
	expect(excel.reduce((sum, { byteLength }) => sum + byteLength, 0)).toBeLessThanOrEqual(limit);
	expect((await store.read('word')).map(({ byteLength }) => byteLength)).toEqual([limit - 1, 1]);
	expect(await store.read('powerpoint')).toEqual([]);
});

it('reopens the database after another connection closes it with a version change', async () => {
	const store = createAddInPreferencesStore(name);
	await store.write('word', 'one', Uint8Array.of(1));
	await new Promise<void>((resolve, reject) => {
		const request = indexedDB.deleteDatabase(name);
		request.onsuccess = () => resolve();
		request.onerror = () => reject(request.error);
	});
	await store.write('word', 'two', Uint8Array.of(2));
	expect(await createAddInPreferencesStore(name).read('word')).toEqual([Uint8Array.of(2)]);
});
