// Per-browser conveniences in localStorage: who I am and the server settings. Every access is
// guarded: storage can be blocked or full, and the app must work without it.
import { type StorageLike, type TeamsServerConfig, parseServerConfig } from 'ooxml-core/teams';

/** localStorage that never throws (private windows, blocked cookies, quota). */
export const safeStorage: StorageLike = {
	getItem(key) {
		try {
			return globalThis.localStorage?.getItem(key) ?? null;
		} catch {
			return null;
		}
	},
	setItem(key, value) {
		try {
			globalThis.localStorage?.setItem(key, value);
		} catch {
			// Blocked or over quota: the convenience is lost, the app keeps working.
		}
	},
	// `storage` events fire in the other tabs of this origin, which is what shares a chosen status.
	subscribe(listener) {
		const target = globalThis as Partial<Pick<Window, 'addEventListener' | 'removeEventListener'>>;
		const onStorage = (event: StorageEvent): void => {
			try {
				if (event.storageArea && event.storageArea !== globalThis.localStorage) return;
			} catch {
				return;
			}
			listener(event.key);
		};
		target.addEventListener?.('storage', onStorage);
		return () => target.removeEventListener?.('storage', onStorage);
	},
};

export interface Identity {
	id: string;
	name: string;
}

export function loadIdentity(): Identity | null {
	try {
		const raw = JSON.parse(
			safeStorage.getItem('teams:identity') ?? 'null',
		) as Partial<Identity> | null;
		if (raw && typeof raw.id === 'string' && typeof raw.name === 'string' && raw.name.trim())
			return { id: raw.id, name: raw.name };
	} catch {
		// fall through
	}
	return null;
}
export const saveIdentity = (identity: Identity): void =>
	safeStorage.setItem('teams:identity', JSON.stringify(identity));

export function loadConfig(): TeamsServerConfig | null {
	const raw = safeStorage.getItem('teams:config');
	if (!raw) return null;
	try {
		return parseServerConfig(JSON.parse(raw)).config;
	} catch {
		return null;
	}
}
export const saveConfig = (config: TeamsServerConfig): void =>
	safeStorage.setItem('teams:config', JSON.stringify(config));
