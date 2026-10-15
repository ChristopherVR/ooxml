// Room access for private chats. A chat's messages live in a sync room of their own,
// `chat-<workspace hash>-<chat id>`, whose Yjs document carries a `teams:chat-members` map
// (user id -> true). A socket may join only when its `user` query parameter is in that map. A room
// that has no members yet is created by its first socket, which must list the members in the
// `members` query parameter and be one of them; a 1:1 chat (`dm-<hash>`) must also list exactly the
// two people its id was derived from, so nobody can create somebody else's 1:1 room first.
// Identity is a claim (there are no accounts): this keeps out workspace members who are not on the
// list, not someone who holds the token and claims another person's id. Pure helpers, no I/O.

export const CHAT_MEMBERS_MAP = 'teams:chat-members';
const CHAT_ROOM = /^chat-[0-9a-f]{16}-([\w-]+)$/u;
const USER = /^[\w.@:+-]{1,64}$/u;
export const MAX_CHAT_MEMBERS = 50;

export const isChatRoom = (name) => CHAT_ROOM.test(name);

/** The same 16-hex hash as `shortHash` in ooxml-core/teams (chats.ts); a test keeps them equal. */
export function shortHash(value) {
	let a = 0x811c9dc5;
	let b = 0x01000193 ^ value.length;
	for (let i = 0; i < value.length; i++) {
		const c = value.charCodeAt(i);
		a = Math.imul(a ^ c, 0x01000193) >>> 0;
		b = Math.imul(b ^ (c + i), 0x85ebca6b) >>> 0;
	}
	return a.toString(16).padStart(8, '0') + b.toString(16).padStart(8, '0');
}

export const directChatId = (a, b) => `dm-${shortHash([a, b].sort().join('\n'))}`;

/** Valid, unique user ids from a comma-separated query parameter. */
export function parseMembers(param) {
	if (typeof param !== 'string' || param.length > 4096) return [];
	return [
		...new Set(
			param
				.split(',')
				.filter((id) => USER.test(id))
				.slice(0, MAX_CHAT_MEMBERS),
		),
	];
}

/** The user ids currently listed in a chat document. */
export function roomMembers(doc) {
	const out = new Set();
	doc.getMap(CHAT_MEMBERS_MAP).forEach((on, id) => {
		if (on === true && USER.test(id)) out.add(id);
	});
	return out;
}

/**
 * Decide whether a socket may join a chat room.
 * `existing` is the set of members already stored for the room (empty for a new room).
 * Returns `{ ok: false }`, or `{ ok: true, seed }` where `seed` lists the members to store for a
 * new room (empty when the room already has members).
 */
export function authorizeChatJoin({ name, user, members, existing }) {
	if (!CHAT_ROOM.test(name) || !USER.test(user ?? '')) return { ok: false };
	if (existing.size > 0) return existing.has(user) ? { ok: true, seed: [] } : { ok: false };
	const listed = parseMembers(members);
	if (!listed.includes(user)) return { ok: false };
	const chatId = CHAT_ROOM.exec(name)[1];
	if (chatId.startsWith('dm-')) {
		if (listed.length !== 2 || directChatId(listed[0], listed[1]) !== chatId) return { ok: false };
	}
	return { ok: true, seed: listed };
}
