const databases = new Map();
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
		void pending.catch(() => databases.delete(name));
	}
	return databases.get(name);
}

function productKeys(host) {
	return IDBKeyRange.bound(`${host}:`, `${host};`, false, true);
}

export function officeAddInPreferences(profile) {
	async function transaction(mode, action) {
		const database = await openAddInPreferences(profile);
		return new Promise((resolve, reject) => {
			const tx = database.transaction('installs', mode);
			const request = action(tx.objectStore('installs'), tx);
			tx.oncomplete = () => resolve(request?.result);
			tx.onabort = () => reject(tx.error ?? new Error('Could not save the add-in preference.'));
			tx.onerror = () => reject(tx.error);
		});
	}
	return {
		read: (host) =>
			transaction('readonly', (store) => store.getAll(productKeys(host), installLimit)),
		write: (host, id, bytes) =>
			transaction('readwrite', (store, tx) => {
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
			}),
	};
}
