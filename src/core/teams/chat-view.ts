// Pure projections for the Chat app: the chat list (recent first, unread, last message preview,
// presence of the other person), and the contacts that "New chat" and a meeting's attendee picker
// offer. No side effects, so every binding gets the same answers. DOM-free. New code.
import { activityPreview } from './activity';
import { type ChatInfo, type ChatKind, chatTitle } from './chats';
import type { Message } from './model';
import {
	type ClientPresence,
	type PresenceActivity,
	type PresenceStatus,
	clientPresenceFromWire,
	combinePresence,
} from './presence';
import type { RosterEntry } from './roster';
import { type PeerLike, callRoomId } from './view';

export interface ChatMemberView {
	id: string;
	name: string;
	self: boolean;
	availability: PresenceStatus;
	/** `meeting` while the status is Busy because of a meeting. */
	activity?: PresenceActivity;
}

export interface ChatLastMessage {
	authorName: string;
	text: string;
	ts: number;
	mine: boolean;
	system: boolean;
}

export interface ChatView {
	id: string;
	kind: ChatKind;
	title: string;
	/** A group's own name, empty for a 1:1 chat or an unnamed group. */
	name: string;
	members: ChatMemberView[];
	unread: number;
	lastMessage: ChatLastMessage | null;
	/** Last activity (last message, else creation); the list is sorted by it. */
	ts: number;
	meetingId?: string;
	/** Someone is in this chat's call right now. */
	live: boolean;
	/** 1:1 chats cannot be left; groups can. */
	canLeave: boolean;
}

export interface ContactView {
	id: string;
	name: string;
	availability: PresenceStatus;
	activity?: PresenceActivity;
}

/**
 * What another person shows as. One person can have several clients (tabs, devices), each its own
 * awareness entry: they are merged by `combinePresence`, never picked by order. Null when none.
 */
export function peerPresenceOf(userId: string, peers: readonly PeerLike[]): ClientPresence | null {
	return combinePresence(
		peers.filter((p) => p.userId === userId).map((p) => clientPresenceFromWire(p)),
	);
}

export function memberPresence(
	userId: string,
	selfId: string,
	selfStatus: PresenceStatus,
	peers: readonly PeerLike[],
): PresenceStatus {
	if (userId === selfId) return selfStatus;
	return peerPresenceOf(userId, peers)?.status ?? 'offline';
}

/** `memberPresence` with the activity behind a derived status. */
function memberShown(
	userId: string,
	selfId: string,
	self: { status: PresenceStatus; activity?: PresenceActivity | undefined },
	peers: readonly PeerLike[],
): { availability: PresenceStatus; activity?: PresenceActivity } {
	const shown =
		userId === selfId
			? { status: self.status, activity: self.activity }
			: (peerPresenceOf(userId, peers) ?? { status: 'offline' as const, activity: undefined });
	return {
		availability: shown.status,
		...(shown.activity ? { activity: shown.activity } : {}),
	};
}

/** Chats the person belongs to, newest activity first. Unread = others' messages after the marker. */
export function chatViews(input: {
	chats: readonly ChatInfo[];
	messagesOf: (chatId: string) => Message[];
	selfId: string;
	selfStatus: PresenceStatus;
	selfActivity?: PresenceActivity;
	lastRead: Record<string, number>;
	peers: readonly PeerLike[];
	workspaceId: string;
	selectedId: string;
	inCallRoom?: string;
}): ChatView[] {
	const { chats, messagesOf, selfId, selfStatus, lastRead, peers, workspaceId, selectedId } = input;
	return chats
		.map((chat): ChatView => {
			const messages = messagesOf(chat.id).filter((m) => !m.deleted);
			const read = lastRead[chat.id] ?? 0;
			const last = messages[messages.length - 1];
			const room = callRoomId(workspaceId, chat.id);
			return {
				id: chat.id,
				kind: chat.kind,
				title: chatTitle(chat, selfId),
				name: chat.name,
				members: chat.members.map((m) => ({
					id: m.id,
					name: m.name,
					self: m.id === selfId,
					...memberShown(m.id, selfId, { status: selfStatus, activity: input.selfActivity }, peers),
				})),
				unread:
					chat.id === selectedId
						? 0
						: messages.filter((m) => !m.system && m.authorId !== selfId && m.ts > read).length,
				lastMessage: last
					? {
							authorName: last.authorName,
							text: activityPreview(last),
							ts: last.ts,
							mine: last.authorId === selfId,
							system: last.system === true,
						}
					: null,
				ts: last?.ts ?? chat.createdAt,
				...(chat.meetingId ? { meetingId: chat.meetingId } : {}),
				live: input.inCallRoom === room || peers.some((p) => p.callRoom === room),
				canLeave: chat.kind === 'group',
			};
		})
		.sort((a, b) => b.ts - a.ts || a.title.localeCompare(b.title));
}

/** Everybody on the roster except you, with presence, online people first. */
export function contactViews(
	roster: readonly RosterEntry[],
	selfId: string,
	peers: readonly PeerLike[],
): ContactView[] {
	const rank = (s: PresenceStatus): number => (s === 'offline' ? 1 : 0);
	return roster
		.filter((r) => r.id !== selfId)
		.map((r) => ({
			id: r.id,
			name: r.name,
			...memberShown(r.id, selfId, { status: 'available' }, peers),
		}))
		.sort((a, b) => rank(a.availability) - rank(b.availability) || a.name.localeCompare(b.name));
}
