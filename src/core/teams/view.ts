// Pure projections from core state to what a UI shows: channel rows with unread and live flags,
// people, typing names, a channel's files and cross-channel message search. No side effects, so
// every binding (and every test) gets the same answers. New code.
import type { CallParticipant } from './call.js';
import type { Availability } from './workspace.js';
import type { Channel, Message, OfficeKind } from './model.js';

export interface ChannelView {
	id: string;
	name: string;
	topic: string;
	kind: 'channel' | 'direct';
	unread: number;
	/** Someone is in a call in this channel right now. */
	live: boolean;
}

/** The slice of a remote presence entry the projections read. */
export interface PeerLike {
	clientId: number;
	userName: string;
	userColor: string;
	channelId?: string;
	typingIn?: string;
	callRoom?: string;
	availability?: string;
}

export interface PersonView {
	id: string;
	name: string;
	color: string;
	self: boolean;
	availability: Availability;
	inCall: boolean;
}

export interface FileEntry {
	name: string;
	kind: OfficeKind;
	size?: number;
	url?: string;
	mime?: string;
	author: string;
	ts: number;
	messageId: string;
	channelId?: string;
	channelName?: string;
}

/** All query words must match the file name, author, channel or recognized kind. */
export function filterFiles(files: readonly FileEntry[], query: string): FileEntry[] {
	const terms = query.trim().toLocaleLowerCase().split(/\s+/u).filter(Boolean);
	return files.filter((file) => {
		const text =
			`${file.name} ${file.author} ${file.channelName ?? ''} ${file.kind}`.toLocaleLowerCase();
		return terms.every((term) => text.includes(term));
	});
}

export interface SearchHit {
	channelId: string;
	channelName: string;
	message: Message;
}

export const callRoomId = (workspaceId: string, channelId: string): string =>
	`call-${workspaceId}-${channelId}`.slice(0, 128);

const asAvailability = (v: unknown): Availability =>
	v === 'busy' || v === 'away' ? v : 'available';

/** Unread = messages by other people newer than the channel's read marker (0 when selected). */
export function channelViews(input: {
	channels: Channel[];
	messagesOf: (channelId: string) => Message[];
	selfId: string;
	lastRead: Record<string, number>;
	peers: PeerLike[];
	workspaceId: string;
	selectedId: string;
	inCallRoom?: string;
}): ChannelView[] {
	const { channels, messagesOf, selfId, lastRead, peers, workspaceId, selectedId, inCallRoom } =
		input;
	return channels
		.filter((c) => !c.archived)
		.map((c) => {
			const read = lastRead[c.id] ?? 0;
			const unread =
				c.id === selectedId
					? 0
					: messagesOf(c.id).filter((m) => !m.deleted && m.authorId !== selfId && m.ts > read)
							.length;
			const room = callRoomId(workspaceId, c.id);
			return {
				id: c.id,
				name: c.name,
				topic: c.topic,
				kind: c.kind,
				unread,
				live: inCallRoom === room || peers.some((p) => p.callRoom === room),
			};
		});
}

export function peopleViews(
	self: { id: string; name: string; color: string; availability: Availability; inCall: boolean },
	peers: PeerLike[],
): PersonView[] {
	return [
		{ ...self, self: true },
		...peers.map((p) => ({
			id: `peer-${p.clientId}`,
			name: p.userName,
			color: p.userColor,
			self: false,
			availability: asAvailability(p.availability),
			inCall: Boolean(p.callRoom),
		})),
	];
}

export const typingNames = (peers: PeerLike[], channelId: string): string[] =>
	peers.filter((p) => p.typingIn === channelId).map((p) => p.userName);

/** Every attachment of a channel, newest first (the Files tab). */
export function filesOf(messages: Message[]): FileEntry[] {
	const out: FileEntry[] = [];
	for (const m of messages) {
		if (m.deleted) continue;
		for (const a of m.attachments)
			out.push({
				name: a.name,
				kind: a.kind,
				...(a.size !== undefined ? { size: a.size } : {}),
				...(a.url ? { url: a.url } : {}),
				...(a.mime ? { mime: a.mime } : {}),
				author: m.authorName,
				ts: m.ts,
				messageId: m.id,
			});
	}
	return out.sort((a, b) => b.ts - a.ts);
}

/** Every attachment in every channel, newest first (the Files view). */
export function allFilesOf(
	channels: Channel[],
	messagesOf: (channelId: string) => Message[],
): FileEntry[] {
	return channels
		.filter((c) => !c.archived)
		.flatMap((c) =>
			filesOf(messagesOf(c.id)).map((f) => ({ ...f, channelId: c.id, channelName: c.name })),
		)
		.sort((a, b) => b.ts - a.ts);
}

/** Case-insensitive message search across channels; newest first, capped. */
export function searchMessages(
	channels: Channel[],
	messagesOf: (channelId: string) => Message[],
	query: string,
	limit = 50,
): SearchHit[] {
	const q = query.trim().toLowerCase();
	if (!q) return [];
	const hits: SearchHit[] = [];
	for (const c of channels) {
		if (c.archived) continue;
		for (const m of messagesOf(c.id))
			if (
				!m.deleted &&
				(m.text.toLowerCase().includes(q) || m.authorName.toLowerCase().includes(q))
			)
				hits.push({ channelId: c.id, channelName: c.name, message: m });
	}
	return hits.sort((a, b) => b.message.ts - a.message.ts).slice(0, limit);
}

export interface CallSummary {
	self: CallParticipant | undefined;
	participants: CallParticipant[];
}
