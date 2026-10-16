/**
 * Version history. Before a write replaces a document's bytes, the bytes it replaces are kept as
 * a version, with what replaced them ("Saved", "Assistant edits", "Restored"). Versions live on
 * this device beside the profile's library, at most `LIMIT` per document, oldest dropped first.
 */
import { currentProfile } from './profile-state.js';
import { $, choose, download, escape, notify, task } from './ui.js';

const LIMIT = 30;
const databases = new Map();

function database() {
	const id = currentProfile().id;
	const name = id === 'default' ? 'ooxml-suite-versions' : `ooxml-suite-versions-${id}`;
	if (!databases.has(name))
		databases.set(
			name,
			new Promise((resolve, reject) => {
				const request = indexedDB.open(name, 1);
				request.onupgradeneeded = () =>
					request.result
						.createObjectStore('versions', { keyPath: 'id' })
						.createIndex('documentId', 'documentId');
				request.onsuccess = () => resolve(request.result);
				request.onerror = () => reject(request.error);
			}),
		);
	return databases.get(name);
}
async function run(mode, action) {
	const db = await database();
	return new Promise((resolve, reject) => {
		const tx = db.transaction('versions', mode);
		const request = action(tx.objectStore('versions'));
		tx.oncomplete = () => resolve(request?.result);
		tx.onerror = () => reject(tx.error);
	});
}
const same = (a, b) => a.length === b.length && a.every((byte, i) => byte === b[i]);

export const versions = {
	async list(documentId) {
		const all = await run('readonly', (s) => s.index('documentId').getAll(documentId));
		return (all ?? []).sort((a, b) => b.created - a.created);
	},
	get: (id) => run('readonly', (s) => s.get(id)),
	/** Keeps `doc` as it is now, labelled with what is about to replace it. */
	async keep(doc, reason) {
		const id = crypto.randomUUID();
		await run('readwrite', (s) =>
			s.put({
				id,
				documentId: doc.id,
				name: doc.name,
				revision: doc.revision,
				modified: doc.modified,
				created: Date.now(),
				reason,
				bytes: doc.bytes,
			}),
		);
		const old = (await this.list(doc.id)).slice(LIMIT);
		if (old.length) await run('readwrite', (s) => old.forEach((v) => s.delete(v.id)));
		return id;
	},
};

/**
 * Wraps the document store so every commit that changes a file's bytes keeps the old bytes as a
 * version first. `reason` describes the write; it defaults to an ordinary save.
 */
export function withVersions(store, enabled = () => true) {
	const commit = store.commit.bind(store);
	store.commit = async (changes, reason = 'Saved') => {
		if (enabled())
			for (const { document } of changes) {
				const current = await store.get(document.id).catch(() => undefined);
				if (current && !same(current.bytes, document.bytes)) await versions.keep(current, reason);
			}
		return commit(changes);
	};
	return store;
}

const when = (time) =>
	new Date(time).toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
const size = (bytes) => `${Math.max(1, Math.round(bytes.length / 1024))} KB`;

/** The version history dialog for one document, with restore and download. */
export async function showVersions({ id, store, sessions, refresh }) {
	const doc = await store.get(id);
	const list = await versions.list(id);
	const content = document.createElement('div');
	content.className = 'version-list';
	content.innerHTML =
		`<p>Earlier versions of <strong>${escape(doc.name)}</strong> kept on this device. Restoring keeps the current version too, so you can always go back.</p>` +
		(list.length
			? `<ol>${list
					.map(
						(v) =>
							`<li><span><strong>${escape(when(v.modified))}</strong><small>Replaced by: ${escape(v.reason)} · ${escape(when(v.created))} · ${size(v.bytes)}</small></span><button type="button" data-version-download="${v.id}">Download</button><button type="button" class="primary" data-version-restore="${v.id}">Restore</button></li>`,
					)
					.join('')}</ol>`
			: '<p class="version-empty">No earlier versions yet. One is kept each time this file is saved or changed by the assistant.</p>');
	content.addEventListener('click', (event) => {
		const restore = event.target.closest('[data-version-restore]')?.dataset.versionRestore;
		const save = event.target.closest('[data-version-download]')?.dataset.versionDownload;
		if (save)
			task(async () => {
				const v = await versions.get(save);
				download({
					name: v.name.replace(/(\.\w+)$/u, ` (${when(v.modified).replace(/[/:]/gu, '-')})$1`),
					bytes: v.bytes,
				});
			});
		if (restore)
			task(async () => {
				const v = await versions.get(restore);
				const current = await sessions.save(id);
				const next = {
					...current,
					bytes: v.bytes,
					revision: current.revision + 1,
					modified: Date.now(),
				};
				await store.commit(
					[{ document: next, expected: current.revision }],
					'Restored an earlier version',
				);
				await sessions.reload(next);
				await refresh();
				$('dialog').close();
				notify(`Restored the version from ${when(v.modified)}`);
			});
	});
	choose('Version history', content);
}
