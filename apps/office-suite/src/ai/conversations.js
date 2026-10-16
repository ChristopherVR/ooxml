/**
 * Chat history, kept on this device in IndexedDB beside the profile's documents. Each profile
 * has its own database, so switching profiles switches the history with the library.
 * A conversation is { id, title, created, updated, pinned, messages: [{ id, role, content, ... }] }.
 */
import { currentProfile } from '../profile-state.js';

const databases = new Map();
const changed = new EventTarget();

function database() {
	const id = currentProfile().id;
	const name = id === 'default' ? 'ooxml-suite-chat' : `ooxml-suite-chat-${id}`;
	if (!databases.has(name))
		databases.set(
			name,
			new Promise((resolve, reject) => {
				const request = indexedDB.open(name, 1);
				request.onupgradeneeded = () =>
					request.result.createObjectStore('conversations', { keyPath: 'id' });
				request.onsuccess = () => resolve(request.result);
				request.onerror = () => reject(request.error);
			}),
		);
	return databases.get(name);
}

async function run(mode, action) {
	const db = await database();
	return new Promise((resolve, reject) => {
		const tx = db.transaction('conversations', mode);
		const request = action(tx.objectStore('conversations'));
		tx.oncomplete = () => resolve(request?.result);
		tx.onerror = () => reject(tx.error);
	});
}

export const conversations = {
	async list() {
		const all = (await run('readonly', (s) => s.getAll())) ?? [];
		return all.sort((a, b) => b.updated - a.updated);
	},
	get: (id) => run('readonly', (s) => s.get(id)),
	async put(conversation) {
		await run('readwrite', (s) => s.put(conversation));
		changed.dispatchEvent(new CustomEvent('change', { detail: conversation.id }));
	},
	async remove(id) {
		await run('readwrite', (s) => s.delete(id));
		changed.dispatchEvent(new CustomEvent('change', { detail: id }));
	},
	onChange(listener) {
		changed.addEventListener('change', listener);
		return () => changed.removeEventListener('change', listener);
	},
};

export const newId = () => crypto.randomUUID();

/** A conversation's title from its first request: one line, at most 60 characters. */
export function titleFor(prompt) {
	const line = prompt.replace(/\s+/gu, ' ').trim();
	return line.length > 60 ? `${line.slice(0, 57).trimEnd()}...` : line || 'New chat';
}

/** History groups in display order, as the chat sidebar lists them. */
export function groupByDate(list, now = Date.now()) {
	const day = 86400000;
	const today = new Date(now).setHours(0, 0, 0, 0);
	const groups = [
		['Pinned', (c) => c.pinned],
		['Today', (c) => c.updated >= today],
		['Yesterday', (c) => c.updated >= today - day],
		['Previous 7 days', (c) => c.updated >= today - 7 * day],
		['Previous 30 days', (c) => c.updated >= today - 30 * day],
		['Older', () => true],
	].map(([label, test]) => ({ label, test, items: [] }));
	for (const c of list) groups.find((g) => g.test(c)).items.push(c);
	return groups.filter((g) => g.items.length);
}
