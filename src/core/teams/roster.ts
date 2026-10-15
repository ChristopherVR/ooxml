// The workspace roster: everybody who has opened the workspace, so "New chat" and a meeting's
// attendee list can offer people who are offline. Each client registers itself in the shared
// document; rows are re-validated when read back. Advisory like presence: there are no accounts.
// DOM-free. New code.
import * as Y from 'yjs';
import { sanitizeChannelName } from './model';
import { type ChatMember, isChatUserId } from './chats';

const ROSTER = 'teams:roster';
const MAX_ROSTER = 500;
const REFRESH_MS = 6 * 60 * 60 * 1000;

export interface RosterEntry extends ChatMember {
	lastSeen: number;
}

export interface Roster {
	/** Add yourself, or refresh your name and last-seen time. Cheap: writes only when needed. */
	register: () => void;
	list: () => RosterEntry[];
}

export function createRoster(doc: Y.Doc, user: ChatMember, now: () => number = Date.now): Roster {
	const map = doc.getMap<Y.Map<unknown>>(ROSTER);
	return {
		register() {
			if (!isChatUserId(user.id)) return;
			const name = sanitizeChannelName(user.name);
			if (!name) return;
			const row = map.get(user.id);
			const seen = row instanceof Y.Map ? Number(row.get('seen')) || 0 : 0;
			if (row instanceof Y.Map && row.get('name') === name && now() - seen < REFRESH_MS) return;
			if (!row && map.size >= MAX_ROSTER) return;
			doc.transact(() => {
				const next = row instanceof Y.Map ? row : new Y.Map<unknown>();
				next.set('name', name);
				next.set('seen', now());
				if (!row) map.set(user.id, next);
			});
		},
		list() {
			const out: RosterEntry[] = [];
			map.forEach((raw, id) => {
				if (!(raw instanceof Y.Map) || !isChatUserId(id)) return;
				const name = sanitizeChannelName(raw.get('name'));
				if (name) out.push({ id, name, lastSeen: Number(raw.get('seen')) || 0 });
			});
			return out.sort((a, b) => a.name.localeCompare(b.name) || a.id.localeCompare(b.id));
		},
	};
}
