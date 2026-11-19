// Private chats (1:1 and group): the directory that lists them. A chat's messages live in a
// document of their own (see `chat-rooms.ts`), so only the people on its member list ever sync
// them. The directory below is the part every workspace member syncs: who is in which chat, so
// each client can find its own chats. Peer-written rows are re-validated when read back, and a
// server with accounts refuses rows the writer may not change (`validate-chats.ts`). DOM-free.
import * as Y from 'yjs';
import {
	CHATS_MAP,
	type ChatInfo,
	type ChatKind,
	type ChatMember,
	MAX_CHAT_MEMBERS,
	createGroupChatId,
	directChatId,
	isChatUserId,
	legacyDirectChatId,
	memberKey,
	readChatRow,
	sanitizeMembers,
} from './chat-records';
import { sanitizeChannelName } from './model';

export {
	type ChatInfo,
	type ChatKind,
	type ChatMember,
	MAX_CHAT_MEMBERS,
	chatRoomId,
	createGroupChatId,
	directChatId,
	isChatUserId,
	isGroupChatIdOf,
	legacyDirectChatId,
	sanitizeMembers,
	shortHash,
} from './chat-records';

/** What the list shows: the group name, or the other people's names. */
export function chatTitle(
	chat: Pick<ChatInfo, 'kind' | 'name' | 'members'>,
	selfId: string,
): string {
	if (chat.name) return chat.name;
	const others = chat.members.filter((m) => m.id !== selfId).map((m) => m.name);
	if (others.length === 0) return 'Just you';
	return chat.kind === 'direct' ? (others[0] ?? 'Chat') : others.slice(0, 4).join(', ');
}

export const isChatMember = (chat: Pick<ChatInfo, 'members'>, userId: string): boolean =>
	chat.members.some((m) => m.id === userId);

export interface ChatDirectory {
	/** Every chat the directory lists (the caller filters to its own). */
	list: () => ChatInfo[];
	get: (id: string) => ChatInfo | null;
	/** Idempotent: the 1:1 chat with `other`, created when missing. Null for yourself or bad ids. */
	ensureDirect: (other: ChatMember) => ChatInfo | null;
	/** A new group of you and `others` (at least one), under an id only you can create. */
	createGroup: (input: {
		others: readonly ChatMember[];
		name?: string;
		meetingId?: string;
	}) => ChatInfo | null;
	/** Members only, groups only. */
	rename: (id: string, name: string) => boolean;
	/** Members only, groups only. Returns the people actually added. */
	add: (id: string, people: readonly ChatMember[]) => ChatMember[];
	/** A member removes someone (or leaves with their own id). Groups only; a group keeps one member. */
	remove: (id: string, userId: string) => boolean;
}

export function createChatDirectory(
	doc: Y.Doc,
	user: ChatMember,
	now: () => number = Date.now,
): ChatDirectory {
	const map = doc.getMap<Y.Map<unknown>>(CHATS_MAP);
	const get = (id: string): ChatInfo | null => readChatRow(id, map.get(id));
	const write = (
		id: string,
		kind: ChatKind,
		name: string,
		members: ChatMember[],
		meetingId?: string,
	): ChatInfo | null => {
		doc.transact(() => {
			const row = new Y.Map<unknown>();
			row.set('kind', kind);
			row.set('name', name);
			row.set('members', members);
			row.set('createdAt', now());
			row.set('createdBy', user.id);
			if (meetingId) row.set('meetingId', meetingId);
			map.set(id, row);
		});
		return get(id);
	};
	const mine = (id: string): { row: Y.Map<unknown>; info: ChatInfo } | null => {
		const info = get(id);
		const row = map.get(id);
		return info && row && isChatMember(info, user.id) ? { row, info } : null;
	};
	return {
		list() {
			const out: ChatInfo[] = [];
			map.forEach((raw, id) => {
				const info = readChatRow(id, raw);
				if (info) out.push(info);
			});
			return out.sort((a, b) => a.createdAt - b.createdAt || a.id.localeCompare(b.id));
		},
		get,
		ensureDirect(other) {
			if (!isChatUserId(user.id) || !isChatUserId(other.id) || other.id === user.id) return null;
			// A chat from before protocol 2 keeps its id, and with it its room and history.
			const existing =
				get(legacyDirectChatId(user.id, other.id)) ?? get(directChatId(user.id, other.id));
			if (existing) return isChatMember(existing, user.id) ? existing : null;
			const members = sanitizeMembers([user, other]);
			return members.length === 2
				? write(directChatId(user.id, other.id), 'direct', '', members)
				: null;
		},
		createGroup({ others, name = '', meetingId }) {
			if (!isChatUserId(user.id)) return null;
			const members = sanitizeMembers([user, ...others]);
			if (members.length < 2) return null;
			const id = createGroupChatId(user.id);
			return write(id, 'group', sanitizeChannelName(name), members, meetingId);
		},
		rename(id, name) {
			const found = mine(id);
			if (!found || found.info.kind !== 'group') return false;
			found.row.set('name', sanitizeChannelName(name));
			return true;
		},
		add(id, people) {
			const found = mine(id);
			if (!found || found.info.kind !== 'group') return [];
			const room = Math.max(0, MAX_CHAT_MEMBERS - found.info.members.length);
			const added = sanitizeMembers(people)
				.filter((m) => !isChatMember(found.info, m.id))
				.slice(0, room);
			doc.transact(() => {
				for (const m of added) found.row.set(memberKey(m.id), m.name);
			});
			return added;
		},
		remove(id, userId) {
			const found = mine(id);
			if (!found || found.info.kind !== 'group' || !isChatMember(found.info, userId)) return false;
			if (found.info.members.length === 1) return false;
			found.row.set(memberKey(userId), null);
			return true;
		},
	};
}
