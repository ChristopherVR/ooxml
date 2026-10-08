import { documentDatabase } from './profile-state.js';
const db = new Promise((resolve, reject) => {
	const request = indexedDB.open(`${documentDatabase()}-locations`, 1);
	request.onupgradeneeded = () => {
		request.result.createObjectStore('folders', { keyPath: 'id' });
		request.result.createObjectStore('sources', { keyPath: 'id' });
	};
	request.onsuccess = () => resolve(request.result);
	request.onerror = () => reject(request.error);
});
export async function locationRecords(store, action, value) {
	const database = await db;
	return new Promise((resolve, reject) => {
		const tx = database.transaction(
			store,
			action === 'getAll' || action === 'get' ? 'readonly' : 'readwrite',
		);
		const request = tx.objectStore(store)[action](...(value === undefined ? [] : [value]));
		tx.oncomplete = () => resolve(request.result);
		tx.onabort = () => reject(tx.error);
		request.onerror = () => reject(request.error);
	});
}
export const desktop = !!window.__TAURI_INTERNALS__;
export async function native(command, args = {}) {
	const { invoke } = await import('@tauri-apps/api/core');
	return invoke(command, args);
}
const snapshots = new Map();
export const forgetSnapshot = (id) => snapshots.delete(id);
export async function chooseFolder() {
	if (desktop) {
		const root = await native('disk_choose_folder');
		return root
			? { id: `native:${root}`, name: root.split(/[\\/]/).pop(), root, type: 'native' }
			: null;
	}
	if (window.showDirectoryPicker) {
		const handle = await window.showDirectoryPicker({ mode: 'read' });
		for (const folder of await locationRecords('folders', 'getAll'))
			if (folder.handle && (await folder.handle.isSameEntry(handle))) return folder;
		return { id: crypto.randomUUID(), name: handle.name, handle, type: 'browser' };
	}
	return new Promise((resolve) => {
		const input = document.createElement('input');
		input.type = 'file';
		input.multiple = true;
		input.setAttribute('webkitdirectory', '');
		input.oncancel = () => resolve(null);
		input.onchange = () => {
			const files = [...input.files];
			if (!files.length) return resolve(null);
			const id = crypto.randomUUID();
			snapshots.set(id, files);
			resolve({
				id,
				name: files[0].webkitRelativePath
					? files[0].webkitRelativePath.split('/')[0]
					: 'Selected files',
				type: 'snapshot',
			});
		};
		input.click();
	});
}
export async function* entries(folder, signal) {
	if (folder.type === 'native') {
		const result = await native('disk_scan', { root: folder.root });
		if (result.truncated)
			folder.warning = 'Search limit reached. Connect a smaller folder to include remaining files.';
		if (result.skipped)
			folder.warning = `${result.skipped} inaccessible folders or files were skipped.`;
		for (const item of result.files) {
			signal.throwIfAborted();
			yield { ...item, folderId: folder.id };
		}
	} else if (folder.type === 'snapshot') {
		for (const file of snapshots.get(folder.id) || []) {
			signal.throwIfAborted();
			yield {
				name: file.name,
				path: file.webkitRelativePath
					? file.webkitRelativePath.split('/').slice(1).join('/')
					: file.name,
				folderId: folder.id,
				file,
			};
		}
	} else {
		let count = 0,
			skipped = 0,
			visited = 0;
		async function* walk(handle, prefix = '', depth = 0) {
			if (depth > 40) {
				skipped++;
				return;
			}
			try {
				for await (const child of handle.values()) {
					signal.throwIfAborted();
					if (++visited > 20000 || count >= 5000) {
						folder.warning =
							'Search limit reached. Connect a smaller folder to include remaining files.';
						return;
					}
					if (child.kind === 'directory') yield* walk(child, prefix + child.name + '/', depth + 1);
					else {
						count++;
						yield {
							name: child.name,
							path: prefix + child.name,
							folderId: folder.id,
							handle: child,
						};
					}
					if (count % 50 === 0) await new Promise((resolve) => setTimeout(resolve, 0));
				}
			} catch (error) {
				if (signal.aborted || error.name === 'NotAllowedError') throw error;
				skipped++;
			}
		}
		yield* walk(folder.handle);
		if (skipped) folder.warning = `${skipped} inaccessible folders were skipped.`;
	}
}
export async function readEntry(folder, entry) {
	if (folder.type === 'native')
		return {
			name: entry.name,
			bytes: new Uint8Array(await native('disk_read', { root: folder.root, path: entry.path })),
		};
	const file = entry.file ?? (await entry.handle.getFile());
	if (file.size > 32 * 1024 * 1024) throw new Error('Choose a file smaller than 32 MB.');
	return { name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) };
}
