import type { WordBuildingBlocksStore } from '../docx/building-blocks-shared';

const STORE = 'values';
const KEY = 'docx-building-blocks';

function openDatabase(name: string): Promise<IDBDatabase | undefined> {
	return new Promise<IDBDatabase | undefined>((resolve) => {
		const request = indexedDB.open(name, 1);
		request.onupgradeneeded = () => request.result.createObjectStore(STORE);
		request.onerror = () => resolve(undefined);
		request.onblocked = () => resolve(undefined);
		request.onsuccess = () => {
			const db = request.result;
			db.onversionchange = () => db.close();
			resolve(db);
		};
	}).catch(() => undefined);
}

/**
 * Word's shared building blocks (Building Blocks.dotx) for one profile, in an IndexedDB database
 * of that name on this device. Without IndexedDB the entries last until the page closes.
 */
export function createBuildingBlocksStore(name: string): WordBuildingBlocksStore {
	let memory: unknown;
	let opening: Promise<IDBDatabase | undefined> | undefined;
	const database = () => {
		const current = (opening ??= openDatabase(name));
		void current.then((db) => {
			if (!db && opening === current) opening = undefined;
			db?.addEventListener('close', () => {
				if (opening === current) opening = undefined;
			});
		});
		return current;
	};
	const run = <T>(
		db: IDBDatabase,
		mode: IDBTransactionMode,
		act: (store: IDBObjectStore) => IDBRequest<T>,
	) =>
		new Promise<T>((resolve, reject) => {
			const tx = db.transaction(STORE, mode);
			const request = act(tx.objectStore(STORE));
			tx.oncomplete = () => resolve(request.result);
			tx.onabort = tx.onerror = () =>
				reject(tx.error ?? new Error('Building blocks store failed.'));
		});
	return {
		async read() {
			const db = typeof indexedDB === 'undefined' ? undefined : await database();
			return db ? run(db, 'readonly', (store) => store.get(KEY)) : structuredClone(memory);
		},
		async write(value) {
			const db = typeof indexedDB === 'undefined' ? undefined : await database();
			if (!db) {
				memory = structuredClone(value);
				return;
			}
			await run(db, 'readwrite', (store) => store.put(value, KEY));
		},
	};
}
