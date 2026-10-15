// Private chats (1:1 and group): the model and the directory that lists them. A chat's messages
// live in a document of their own (see `chat-rooms.ts`), so only the people on its member list
// ever sync them. The directory below is the part every workspace member syncs: who is in which
// chat, so each client can find its own chats. Peer-written rows are re-validated when read back.
// DOM-free. New code.
import * as Y from 'yjs';
import { createIdGenerator } from '../collab/identity';
import { sanitizeChannelName } from './model';

export type ChatKind = 'direct' | 'group';

export interface ChatMember {
	id: string;
	name: string;
}

export interface ChatInfo {
	id: string;
	kind: ChatKind;
	/** The group's name; empty for 1:1 chats and unnamed groups. */
	name: string;
	members: ChatMember[];
	createdAt: number;
	createdBy: string;
	/** Set when the chat belongs to a scheduled meeting. */
	meetingId?: string;
}

export const MAX_CHAT_MEMBERS = 50;
const CHATS = 'teams:chats';
const SAFE_USER = /^[\w.@:+-]{1,64}$/u;
const SAFE_CHAT = /^[\w-]{1,64}$/u;

/** Ids the server can carry in a query string; people with other ids cannot be chatted with. */
export const isChatUserId = (id: unknown): id is string =>
	typeof id === 'string' && SAFE_USER.test(id);

/** Two 32-bit FNV-1a hashes as 16 hex digits: a stable short name, not a security measure. */
export function shortHash(value: string): string {
	let a = 0x811c9dc5;
	let b = 0x01000193 ^ value.length;
	for (let i = 0; i < value.length; i++) {
		const c = value.charCodeAt(i);
		a = Math.imul(a ^ c, 0x01000193) >>> 0;
		b = Math.imul(b ^ (c + i), 0x85ebca6b) >>> 0;
	}
	return a.toString(16).padStart(8, '0') + b.toString(16).padStart(8, '0');
}

/** The same chat for the same two people, whoever opens it first. */
export const directChatId = (a: string, b: string): string =>
	`dm-${shortHash([a, b].sort().join('\n'))}`;

/** The sync room of a chat. At most 128 characters, whatever the workspace id. */
export const chatRoomId = (workspaceId: string, chatId: string): string =>
	`chat-${shortHash(workspaceId)}-${chatId}`.slice(0, 128);

export function sanitizeMembers(raw: unknown): ChatMember[] {
	if (!Array.isArray(raw)) return [];
	const seen = new Set<string>();
	const out: ChatMember[] = [];
	for (const item of raw) {
		if (!item || typeof item !== 'object') continue;
		const { id, name } = item as Record<string, unknown>;
		if (!isChatUserId(id) || seen.has(id)) continue;
		seen.add(id);
		out.push({ id, name: sanitizeChannelName(name) || 'Unknown' });
		if (out.length >= MAX_CHAT_MEMBERS) break;
	}
	return out;
}

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

function readChat(id: string, raw: unknown): ChatInfo | null {
	if (!(raw instanceof Y.Map) || !SAFE_CHAT.test(id)) return null;
	const members = sanitizeMembers(raw.get('members'));
	if (members.length === 0) return null;
	const kind: ChatKind = raw.get('kind') === 'group' ? 'group' : 'direct';
	if (kind === 'direct' && members.length > 2) return null;
	const meetingId = raw.get('meetingId');
	return {
		id,
		kind,
		name: kind === 'group' ? sanitizeChannelName(raw.get('name')) : '',
		members,
		createdAt: Number(raw.get('createdAt')) || 0,
		createdBy: String(raw.get('createdBy') ?? '').slice(0, 160),
		...(typeof meetingId === 'string' && SAFE_CHAT.test(meetingId) ? { meetingId } : {}),
	};
}

export interface ChatDirectory {
	/** Every chat the directory lists (the caller filters to its own). */
	list: () => ChatInfo[];
	get: (id: string) => ChatInfo | null;
	/** Idempotent: the 1:1 chat with `other`, created when missing. Null for yourself or bad ids. */
	ensureDirect: (other: ChatMember) => ChatInfo | null;
	/** A group of you and `others` (at least one). Idempotent for a given `id`. */
	createGroup: (input: {
		others: readonly ChatMember[];
		name?: string;
		id?: string;
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
	const map = doc.getMap<Y.Map<unknown>>(CHATS);
	const nextId = createIdGenerator(String(doc.clientID), 'g');
	const get = (id: string): ChatInfo | null => readChat(id, map.get(id));
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
				const info = readChat(id, raw);
				if (info) out.push(info);
			});
			return out.sort((a, b) => a.createdAt - b.createdAt || a.id.localeCompare(b.id));
		},
		get,
		ensureDirect(other) {
			if (!isChatUserId(user.id) || !isChatUserId(other.id) || other.id === user.id) return null;
			const id = directChatId(user.id, other.id);
			const existing = get(id);
			if (existing) return isChatMember(existing, user.id) ? existing : null;
			const members = sanitizeMembers([user, other]);
			return members.length === 2 ? write(id, 'direct', '', members) : null;
		},
		createGroup({ others, name = '', id, meetingId }) {
			if (!isChatUserId(user.id)) return null;
			const members = sanitizeMembers([user, ...others]);
			if (members.length < 2) return null;
			const chatId = id && SAFE_CHAT.test(id) ? id : nextId('g').replace(/[^\w-]/gu, '');
			const existing = get(chatId);
			if (existing) return isChatMember(existing, user.id) ? existing : null;
			return write(chatId, 'group', sanitizeChannelName(name), members, meetingId);
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
			const merged = sanitizeMembers([...found.info.members, ...people]);
			const added = merged.filter((m) => !isChatMember(found.info, m.id));
			if (added.length) found.row.set('members', merged);
			return added;
		},
		remove(id, userId) {
			const found = mine(id);
			if (!found || found.info.kind !== 'group' || !isChatMember(found.info, userId)) return false;
			const rest = found.info.members.filter((m) => m.id !== userId);
			if (rest.length === 0) return false;
			found.row.set('members', rest);
			return true;
		},
	};
}
