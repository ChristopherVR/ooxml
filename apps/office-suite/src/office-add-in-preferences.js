const databases = new Map();
const sessionStores = new Map();
const installLimit = 32;
const productByteLimit = 16 * 1024 * 1024;

function overLimit(count, bytes) {
	return count > installLimit || bytes > productByteLimit;
}

function openAddInPreferences(profile) {
	const name = `ooxml-suite-add-ins-${profile}`;
	if (!databases.has(name)) {
		const pending = new Promise((resolve, reject) => {
			const request = indexedDB.open(name, 1);
			request.onupgradeneeded = () => request.result.createObjectStore('installs');
			request.onsuccess = () => resolve(request.result);
			request.onerror = () => reject(request.error);
		});
		databases.set(name, pending);
	}
	return databases.get(name);
}

function productKeys(host) {
	return IDBKeyRange.bound(`${host}:`, `${host};`, false, true);
}

function sessionAddInPreferences(profile) {
	if (!sessionStores.has(profile)) sessionStores.set(profile, new Map());
	const products = sessionStores.get(profile);
	function records(host) {
		if (!products.has(host)) products.set(host, new Map());
		return products.get(host);
	}
	return {
		read: (host) => [...records(host).values()].map((bytes) => bytes.slice()),
		write(host, id, bytes) {
			const entries = records(host);
			if (bytes === undefined) entries.delete(id);
			else {
				let total = bytes.byteLength;
				for (const [key, value] of entries) if (key !== id) total += value.byteLength;
				const count = entries.size + Number(!entries.has(id));
				if (overLimit(count, total)) throw new Error('Add-in preference limit reached.');
				entries.set(id, bytes.slice());
			}
		},
	};
}

export function officeAddInPreferences(profile) {
	const fallback = sessionAddInPreferences(profile);
	const database = () => openAddInPreferences(profile).catch(() => undefined);
	function transaction(database, mode, action) {
		return new Promise((resolve, reject) => {
			const tx = database.transaction('installs', mode);
			action(tx.objectStore('installs'), tx, (value) => {
				tx.oncomplete = () => resolve(value);
			});
			tx.onabort = () => reject(tx.error ?? new Error('Could not save the add-in preference.'));
			tx.onerror = () => reject(tx.error);
		});
	}
	return {
		async read(host) {
			const db = await database();
			return db
				? transaction(db, 'readonly', (store, _tx, finish) => {
						const files = [];
						let total = 0;
						const request = store.openCursor(productKeys(host));
						request.onsuccess = () => {
							const cursor = request.result;
							if (!cursor || files.length >= installLimit) return finish(files);
							const bytes = cursor.value;
							if (
								bytes instanceof Uint8Array &&
								!overLimit(files.length + 1, total + bytes.byteLength)
							) {
								files.push(bytes);
								total += bytes.byteLength;
							}
							cursor.continue();
						};
					})
				: fallback.read(host);
		},
		async write(host, id, bytes) {
			const db = await database();
			if (!db) return fallback.write(host, id, bytes);
			return transaction(db, 'readwrite', (store, tx, finish) => {
				finish(undefined);
				const key = `${host}:${id}`;
				if (bytes === undefined) {
					store.delete(key);
					return;
				}
				let count = 1;
				let total = bytes.byteLength;
				const request = store.openCursor(productKeys(host));
				request.onsuccess = () => {
					const cursor = request.result;
					if (overLimit(count, total)) tx.abort();
					else if (!cursor) store.put(bytes, key);
					else {
						if (cursor.key !== key) {
							count++;
							total += cursor.value?.byteLength ?? 0;
						}
						cursor.continue();
					}
				};
			});
		},
	};
}
