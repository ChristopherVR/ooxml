import type { OfficeAddInPreferences } from '../ribbon/add-in-preferences';

const installLimit = 32;
const byteLimit = 16 * 1024 * 1024;

const overLimit = (count: number, bytes: number) => count > installLimit || bytes > byteLimit;
const limitError = () => new Error('Add-in preference limit reached.');

function openDatabase(name: string, closed: () => void): Promise<IDBDatabase | undefined> {
	return new Promise<IDBDatabase | undefined>((resolve) => {
		const request = indexedDB.open(name, 1);
		let blocked = false;
		request.onupgradeneeded = () => request.result.createObjectStore('installs');
		request.onerror = () => resolve(undefined);
		request.onblocked = () => {
			blocked = true;
			resolve(undefined);
		};
		request.onsuccess = () => {
			const db = request.result;
			if (blocked) db.close();
			else {
				// A newer version elsewhere closes this connection; the next request reopens it.
				db.onversionchange = () => {
					db.close();
					closed();
				};
				db.onclose = closed;
				resolve(db);
			}
		};
	}).catch(() => undefined);
}

function transaction<T>(
	db: IDBDatabase,
	mode: IDBTransactionMode,
	run: (store: IDBObjectStore, tx: IDBTransaction) => T,
): Promise<T> {
	return new Promise((resolve, reject) => {
		const tx = db.transaction('installs', mode);
		tx.onabort = () => reject(tx.error ?? limitError());
		const result = run(tx.objectStore('installs'), tx);
		tx.oncomplete = () => resolve(result);
	});
}

/** Every key of one product: `${host}:` up to the next character after the colon. */
const productRange = (host: string) => IDBKeyRange.bound(`${host}:`, `${host};`, false, true);

/** The IndexedDB implementation behind `createAddInPreferencesStore`. */
export function createAddInPreferencesStore(name: string): OfficeAddInPreferences {
	const memory = new Map<string, Uint8Array>();
	let opening: Promise<IDBDatabase | undefined> | undefined;
	const database = () => {
		const current = (opening ??= openDatabase(name, () => {
			if (opening === current) opening = undefined;
		}));
		return current;
	};
	return {
		async read(host) {
			const db = await database();
			if (!db) {
				return [...memory]
					.filter(([key]) => key.startsWith(`${host}:`))
					.map(([, bytes]) => bytes.slice());
			}
			return transaction(db, 'readonly', (store) => {
				const files: Uint8Array[] = [];
				let total = 0;
				// Each product reads only its own records under the limits, so an older database
				// written with per-product limits cannot let one product's add-ins hide another's.
				const request = store.openCursor(productRange(host));
				request.onsuccess = () => {
					const cursor = request.result;
					if (!cursor || files.length >= installLimit || total >= byteLimit) return;
					const bytes: unknown = cursor.value;
					if (bytes instanceof Uint8Array && total + bytes.byteLength <= byteLimit) {
						total += bytes.byteLength;
						files.push(bytes);
					}
					cursor.continue();
				};
				return files;
			});
		},
		async write(host, id, bytes) {
			const db = await database();
			const key = `${host}:${id}`;
			let count = 1;
			let total = bytes?.byteLength ?? 0;
			if (!db) {
				if (!bytes) {
					memory.delete(key);
					return;
				}
				for (const [record, value] of memory) {
					if (record === key) continue;
					if (record.startsWith(`${host}:`)) count++;
					total += value.byteLength;
				}
				if (overLimit(count, total)) throw limitError();
				memory.set(key, bytes.slice());
				return;
			}
			return transaction(db, 'readwrite', (store, tx) => {
				if (!bytes) {
					store.delete(key);
					return;
				}
				const request = store.openCursor();
				request.onsuccess = () => {
					const cursor = request.result;
					if (overLimit(count, total)) tx.abort();
					else if (!cursor) store.put(bytes, key);
					else {
						const value: unknown = cursor.value;
						if (cursor.key !== key && value instanceof Uint8Array) {
							if (typeof cursor.key === 'string' && cursor.key.startsWith(`${host}:`)) count++;
							total += value.byteLength;
						}
						cursor.continue();
					}
				};
			});
		},
	};
}
