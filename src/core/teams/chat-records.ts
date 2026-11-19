// What a private chat is on the wire: its ids, its room name and its row in the directory. The
// client, the update validator and the reference server (through the generated validator bundle)
// all read these, so an id one of them accepts means the same thing to the others. DOM-free.
import * as Y from 'yjs';
import { sha256 } from '../digest/sha256';
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
export const CHATS_MAP = 'teams:chats';
const MEMBER_KEY = 'member:';
const SAFE_USER = /^[\w.@:+-]{1,64}$/u;
const SAFE_CHAT = /^[\w-]{1,64}$/u;
const GROUP_CHAT = /^g-([0-9a-f]{32})-([0-9a-f]{24})$/u;

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

const toHex = (bytes: Uint8Array): string =>
	Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
const digestHex = (text: string, bytes: number): string =>
	toHex(sha256(new TextEncoder().encode(text)).subarray(0, bytes));

/**
 * The same chat for the same two people, whoever opens it first. Collision resistant: the server
 * creates a 1:1 room only for the two people this id was derived from.
 */
export const directChatId = (a: string, b: string): string =>
	`dm-${digestHex(`openteams-direct\n${[a, b].sort().join('\n')}`, 16)}`;

/** The id 1:1 chats had before protocol 2. Their rooms already exist; none is created any more. */
export const legacyDirectChatId = (a: string, b: string): string =>
	`dm-${shortHash([a, b].sort().join('\n'))}`;

const groupTag = (nonce: string, creator: string): string =>
	digestHex(`openteams-group\n${nonce}\n${creator}`, 12);

/** A new group chat id: 128 random bits, bound to the one person allowed to create its room. */
export function createGroupChatId(creator: string): string {
	const nonce = toHex(globalThis.crypto.getRandomValues(new Uint8Array(16)));
	return `g-${nonce}-${groupTag(nonce, creator)}`;
}

/** Whether `id` is a group chat id made by {@link createGroupChatId} for `creator`. */
export function isGroupChatIdOf(id: string, creator: string): boolean {
	const match = GROUP_CHAT.exec(id);
	return match !== null && match[2] === groupTag(match[1]!, creator);
}

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

/**
 * The row field holding one person's place in a group: their name, or null once they are out. A
 * group row keeps the `members` it was created with and never rewrites them, so people added and
 * leaving at the same time touch different fields and Yjs merges them.
 */
export const memberKey = (userId: string): string => `${MEMBER_KEY}${userId}`;

/** The user id of a {@link memberKey} field, or null for any other field. */
export const userOfMemberKey = (key: string): string | null =>
	key.startsWith(MEMBER_KEY) ? key.slice(MEMBER_KEY.length) : null;

/** Whether someone is in a group row now, was in it once, or never was. */
export function memberStanding(
	row: Y.Map<unknown>,
	userId: string,
): 'member' | 'former' | 'stranger' {
	const entry = row.get(memberKey(userId));
	if (entry !== undefined) return typeof entry === 'string' ? 'member' : 'former';
	return sanitizeMembers(row.get('members')).some((m) => m.id === userId) ? 'member' : 'stranger';
}

/** The people of a group row, in an order every client computes alike. */
function groupMembers(row: Y.Map<unknown>): ChatMember[] {
	const members = new Map(sanitizeMembers(row.get('members')).map((m) => [m.id, m]));
	for (const key of [...row.keys()].sort()) {
		const id = userOfMemberKey(key);
		if (id === null || !isChatUserId(id)) continue;
		const name = row.get(key);
		if (typeof name === 'string')
			members.set(id, { id, name: sanitizeChannelName(name) || 'Unknown' });
		else members.delete(id);
	}
	return [...members.values()].slice(0, MAX_CHAT_MEMBERS);
}

const isDirectRow = (id: string, members: readonly ChatMember[]): boolean => {
	const [a, b] = members;
	return (
		members.length === 2 &&
		a !== undefined &&
		b !== undefined &&
		(id === directChatId(a.id, b.id) || id === legacyDirectChatId(a.id, b.id))
	);
};

/** A directory row as the chat it stands for, or null when it is not a well-formed chat. */
export function readChatRow(id: string, raw: unknown): ChatInfo | null {
	if (!(raw instanceof Y.Map) || !SAFE_CHAT.test(id)) return null;
	const kind = raw.get('kind');
	if (kind !== 'direct' && kind !== 'group') return null;
	const members = kind === 'group' ? groupMembers(raw) : sanitizeMembers(raw.get('members'));
	if (members.length === 0 || (kind === 'direct' && !isDirectRow(id, members))) return null;
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
