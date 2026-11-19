// Room access for private chats. A chat's messages live in a sync room of their own,
// `chat-<workspace hash>-<chat id>`, whose Yjs document carries a `teams:chat-members` map
// (user id -> true). A socket may join only when its user is `true` in that map. Nobody names the
// members of a new room: a 1:1 room is created for the two people its id was derived from, a group
// room for the one person its id is bound to, who then adds the others through validated updates.
// A room that already has content never starts again, even when its member list is empty.
// Without accounts the user is a claim: this keeps out workspace members who are not on the list,
// not someone who holds the token and claims another person's id. Pure helpers, no I/O.
import {
	chatIdOfRoom,
	directChatId,
	isChatRoom,
	isChatRoomAlias,
	isCurrentChatId,
	isGroupChatIdOf,
	shortHash,
} from './validator.generated.mjs';

export { chatIdOfRoom, directChatId, isChatRoom, isChatRoomAlias, isCurrentChatId, shortHash };

export const CHAT_MEMBERS_MAP = 'teams:chat-members';
const USER = /^[\w.@:+-]{1,64}$/u;
const DIRECT_CHAT = /^dm-[0-9a-f]{32}$/u;

/** The two valid, distinct user ids of a `members` query parameter, or null. */
function parsePair(param) {
	if (typeof param !== 'string' || param.length > 256) return null;
	const ids = param.split(',');
	return ids.length === 2 && ids[0] !== ids[1] && ids.every((id) => USER.test(id)) ? ids : null;
}

/** The user ids currently listed in a chat document. */
export function roomMembers(doc) {
	const out = new Set();
	doc.getMap(CHAT_MEMBERS_MAP).forEach((on, id) => {
		if (on === true && USER.test(id)) out.add(id);
	});
	return out;
}

/** Who a new room starts with when `user` opens it, or null when they may not create it. */
function newRoomMembers(chatId, user, members) {
	if (isGroupChatIdOf(chatId, user)) return [user];
	if (!DIRECT_CHAT.test(chatId)) return null;
	const pair = parsePair(members);
	return pair?.includes(user) && directChatId(pair[0], pair[1]) === chatId ? pair : null;
}

/**
 * Decide whether a socket may join a chat room. `existing` is the set of members stored for the
 * room and `started` whether the room holds anything at all. Returns `{ ok: false }`, or
 * `{ ok: true, seed }` where `seed` lists the members to store for a new room.
 */
export function authorizeChatJoin({ name, user, members, existing, started }) {
	const chatId = chatIdOfRoom(name);
	if (chatId === null || !USER.test(user ?? '')) return { ok: false };
	if (existing.has(user)) return { ok: true, seed: [] };
	if (started || existing.size > 0) return { ok: false };
	const seed = newRoomMembers(chatId, user, members);
	return seed ? { ok: true, seed } : { ok: false };
}
