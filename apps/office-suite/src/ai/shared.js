/**
 * Shared chats: a conversation that lives in a collaboration room (ooxml-core/teams shared-chat),
 * so people in other workspaces, profiles and sessions work in the same chat. The room syncs over
 * the server the policy names, else the Teams server the user connected, else the tabs of this
 * browser. Each message keeps the suite's own fields in its `data`, which are re-validated on use.
 */
import { currentProfile } from '../profile-state.js';
import { currentPolicy } from './policy.js';

const rooms = new Map();
let core;
const teamsCore = () => (core ??= import('ooxml-core/teams'));

/** The server shared chats use, and a sentence saying who can join. */
export async function collaborationTarget() {
	const { parseServerConfig, localServerConfig } = await teamsCore();
	const { syncUrl, token } = currentPolicy().collaboration;
	let config = localServerConfig();
	let source = 'browser';
	if (syncUrl) {
		config = parseServerConfig({ mode: 'server', syncUrl, ...(token ? { token } : {}) }).config;
		source = 'organisation';
	} else
		try {
			const saved = parseServerConfig(JSON.parse(localStorage.getItem('teams:config'))).config;
			if (saved.mode === 'server' && saved.syncUrl) {
				config = saved;
				source = 'teams';
			}
		} catch {}
	const host = config.syncUrl ? new URL(config.syncUrl).host : '';
	return {
		config,
		source,
		reach:
			source === 'browser'
				? 'Only other tabs of this browser can join, because no sync server is connected. Connect Teams to your server (or ask your administrator for one) to chat with other people.'
				: `Anyone with the link who can reach ${host} can join, from any workspace.`,
	};
}

/** Joins (or reuses) the room of a shared chat. */
export async function joinRoom(room) {
	if (rooms.has(room)) return rooms.get(room);
	const { createSharedChat } = await teamsCore();
	const { config } = await collaborationTarget();
	const profile = currentProfile();
	const chat = createSharedChat({
		roomId: room,
		user: { id: profile.userId, name: profile.name },
		config,
	});
	rooms.set(room, chat);
	return chat;
}

export async function newRoomId() {
	return (await teamsCore()).sharedChatRoomId();
}
export async function isRoomId(value) {
	return (await teamsCore()).isSharedChatRoomId(value);
}

/** The link that opens a shared chat in the suite. */
export const inviteLink = (room) => {
	const url = new URL(location.href);
	url.hash = `#/chat/join/${room}`;
	return url.href;
};
/** A room id from a pasted link or the bare id. */
export function roomFromInvite(text) {
	return /(chat-[0-9a-f]{32})\s*$/u.exec(String(text).trim())?.[1] ?? null;
}

const OWN = ['id', 'role', 'content', 'author', 'streaming'];

/** One suite message as a shared row. */
export function toShared(message) {
	return {
		id: message.id,
		role: message.role,
		content: message.content ?? '',
		streaming: Boolean(message.streaming),
		data: Object.fromEntries(Object.entries(message).filter(([k]) => !OWN.includes(k))),
	};
}
/** One shared row as a suite message. */
export function fromShared(row) {
	return {
		...row.data,
		id: row.id,
		role: row.role,
		content: row.content,
		author: { id: row.authorId, name: row.authorName },
		...(row.streaming ? { streaming: true } : {}),
	};
}
