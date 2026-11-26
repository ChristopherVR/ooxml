const databases = new Map();
const sessionStores = new Map();
const installLimit = 32;

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
			else if (entries.has(id) || entries.size < installLimit) entries.set(id, bytes.slice());
			else throw new Error('Could not save the add-in preference.');
		},
	};
}

export function officeAddInPreferences(profile) {
	const fallback = sessionAddInPreferences(profile);
	const database = () => openAddInPreferences(profile).catch(() => undefined);
	function transaction(database, mode, action) {
		return new Promise((resolve, reject) => {
			const tx = database.transaction('installs', mode);
			const request = action(tx.objectStore('installs'), tx);
			tx.oncomplete = () => resolve(request?.result);
			tx.onabort = () => reject(tx.error ?? new Error('Could not save the add-in preference.'));
			tx.onerror = () => reject(tx.error);
		});
	}
	return {
		async read(host) {
			const db = await database();
			return db
				? transaction(db, 'readonly', (store) => store.getAll(productKeys(host), installLimit))
				: fallback.read(host);
		},
		async write(host, id, bytes) {
			const db = await database();
			if (!db) return fallback.write(host, id, bytes);
			return transaction(db, 'readwrite', (store, tx) => {
				const key = `${host}:${id}`;
				if (bytes === undefined) return store.delete(key);
				const current = store.getKey(key);
				current.onsuccess = () => {
					if (current.result !== undefined) {
						store.put(bytes, key);
						return;
					}
					const count = store.count(productKeys(host));
					count.onsuccess = () => {
						if (count.result < installLimit) store.put(bytes, key);
						else tx.abort();
					};
				};
			});
		},
	};
}
