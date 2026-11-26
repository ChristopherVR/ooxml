const countLimit = 32;
const byteLimit = 16 * 1024 * 1024;

export function demoAddInPreferences() {
	if (new URLSearchParams(location.search).get('addInsStorage') === 'off') return null;
	const memory = new Map<string, Uint8Array>();
	let opening: Promise<IDBDatabase | undefined> | undefined;
	const database = () =>
		(opening ??= new Promise<IDBDatabase | undefined>((resolve) => {
			const request = indexedDB.open('ooxml-demo-add-ins', 1);
			request.onupgradeneeded = () => request.result.createObjectStore('installs');
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => resolve(undefined);
			request.onblocked = () => resolve(undefined);
		}).catch(() => undefined));
	const range = (product: string) => IDBKeyRange.bound(`${product}:`, `${product};`, false, true);
	const key = (product: string, id: string) => `${product}:${id}`;
	const within = (files: readonly Uint8Array[], added: Uint8Array) =>
		files.length < countLimit &&
		files.reduce((total, file) => total + file.byteLength, added.byteLength) <= byteLimit;
	return {
		async read(product: string): Promise<readonly Uint8Array[]> {
			const db = await database();
			if (!db)
				return [...memory]
					.filter(([id]) => id.startsWith(`${product}:`))
					.map(([, bytes]) => bytes.slice());
			return new Promise((resolve, reject) => {
				const request = db
					.transaction('installs')
					.objectStore('installs')
					.getAll(range(product), countLimit);
				request.onsuccess = () =>
					resolve(
						request.result.filter((entry): entry is Uint8Array => entry instanceof Uint8Array),
					);
				request.onerror = () => reject(request.error);
			});
		},
		async write(product: string, id: string, bytes: Uint8Array | undefined): Promise<void> {
			const db = await database();
			const record = key(product, id);
			if (!db) {
				if (!bytes) {
					memory.delete(record);
					return;
				}
				const files = [...memory]
					.filter(([key]) => key !== record && key.startsWith(`${product}:`))
					.map(([, value]) => value);
				if (!within(files, bytes)) throw new Error('Add-in preference limit reached.');
				memory.set(record, bytes.slice());
				return;
			}
			return new Promise((resolve, reject) => {
				const tx = db.transaction('installs', 'readwrite');
				const store = tx.objectStore('installs');
				tx.oncomplete = () => resolve();
				tx.onabort = () => reject(tx.error ?? new Error('Add-in preference limit reached.'));
				if (!bytes) {
					store.delete(record);
					return;
				}
				const cursor = store.openCursor(range(product));
				let size = bytes.byteLength;
				let count = 1;
				cursor.onsuccess = () => {
					if (size > byteLimit || count > countLimit) tx.abort();
					else if (!cursor.result) store.put(bytes, record);
					else {
						if (cursor.result.key !== record) {
							size += cursor.result.value.byteLength;
							count++;
						}
						cursor.result.continue();
					}
				};
			});
		},
	};
}
