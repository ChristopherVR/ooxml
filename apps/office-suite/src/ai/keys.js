/**
 * Keys a user enters for a model. They are kept in this tab's session storage only, so they are
 * gone when the tab closes, and no other page on a shared origin can find them in local storage.
 */
import { currentProfile } from '../profile-state.js';

const PREFIX = 'ooxml-suite-ai-keys-v2-';
const mapKey = () => `${PREFIX}${currentProfile().id}`;

function read() {
	try {
		return JSON.parse(sessionStorage.getItem(mapKey())) ?? {};
	} catch {
		return {};
	}
}

export const keyFor = (id) => read()[id] ?? '';

/** Sets or, with an empty key, forgets the key of one connection. */
export function storeKey(id, key) {
	const keys = read();
	delete keys[id];
	if (key) keys[id] = key;
	try {
		sessionStorage.setItem(mapKey(), JSON.stringify(keys));
	} catch {}
}

/**
 * Deletes the keys earlier versions could remember in local storage: the key map of every profile,
 * and the key in a first-version connection record. Run once when the app loads.
 */
export function forgetRememberedKeys() {
	let names = [];
	try {
		names = Array.from({ length: localStorage.length }, (_, i) => localStorage.key(i));
	} catch {}
	for (const name of names)
		try {
			if (name?.startsWith(PREFIX)) localStorage.removeItem(name);
			else if (name?.startsWith('ooxml-suite-ai-v1-')) {
				const { key, remember, ...rest } = JSON.parse(localStorage.getItem(name)) ?? {};
				if (key !== undefined || remember !== undefined)
					localStorage.setItem(name, JSON.stringify(rest));
			}
		} catch {
			localStorage.removeItem(name);
		}
}
