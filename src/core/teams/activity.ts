// The activity feed, derived from the shared chat data: mentions of you, replies to your messages
// and to threads you follow, reactions to your messages and pinned messages. Pure functions: the
// feed is a projection (nothing is stored in Yjs for it), so every client computes the same items
// from the same document, and read state is a personal overlay (`activity-state.ts`). DOM-free.
import type { Channel, Message } from './model';
import { mentionDisplayText } from './mention-text';
import { mentionsUser } from './mentions';
import type { PinnedMessage } from './pins';
import { channelThreads } from './threads';

export type ActivityKind = 'mention' | 'reply' | 'reaction' | 'pin' | 'meeting';

export interface ActivityItem {
	/** Stable across recomputation, so read state and notifications can refer to it. */
	id: string;
	kind: ActivityKind;
	channelId: string;
	channelName: string;
	/** The message to scroll to: the mentioning or replying message, or the one reacted to or pinned. */
	messageId: string;
	/** The thread root of that message (itself when it is a post). */
	rootId: string;
	actorId: string;
	actorName: string;
	ts: number;
	/** A one-line excerpt of the message. */
	preview: string;
	/** The emoji, for reactions. */
	emoji?: string;
	/** The conversation is a private chat rather than a channel. */
	chat?: boolean;
	read: boolean;
}

export interface ReactionLike {
	messageId: string;
	emoji: string;
	userId: string;
	ts: number;
}

export interface ActivityInput {
	selfId: string;
	channels: readonly Channel[];
	messagesOf: (channelId: string) => Message[];
	reactions: readonly ReactionLike[];
	pinsOf: (channelId: string) => PinnedMessage[];
	/** Whether the person follows the thread rooted at this message. */
	isFollowed?: (channelId: string, rootId: string) => boolean;
	/** Read state; items are unread when omitted. */
	isRead?: (id: string, ts: number) => boolean;
}

export const MAX_ACTIVITY_ITEMS = 200;
const PREVIEW_CHARS = 140;

export const activityPreview = (
	message: Pick<Message, 'text' | 'attachments' | 'mentions'>,
): string => {
	const text = mentionDisplayText(message).replace(/\s+/gu, ' ').trim();
	if (text) return text.length > PREVIEW_CHARS ? `${text.slice(0, PREVIEW_CHARS)}...` : text;
	return message.attachments[0]?.name ?? 'Attachment';
};

/** Newest first. Your own actions never appear; deleted messages and archived channels are skipped. */
export function deriveActivity(input: ActivityInput): ActivityItem[] {
	const { selfId } = input;
	const isRead = input.isRead ?? (() => false);
	const items: ActivityItem[] = [];
	const located = new Map<string, { message: Message; channel: Channel; rootId: string }>();
	// Reactions carry only an id; a person is known by the name on their messages.
	const names = new Map<string, string>();
	for (const channel of input.channels) {
		if (channel.archived) continue;
		const messages = input.messagesOf(channel.id);
		const threads = channelThreads(messages);
		for (const message of messages) {
			const rootId = threads.thread(message.id)?.root.id ?? message.id;
			located.set(message.id, { message, channel, rootId });
			names.set(message.authorId, message.authorName);
			if (message.deleted || message.authorId === selfId) continue;
			const base = {
				...(channel.kind === 'direct' ? { chat: true } : {}),
				channelId: channel.id,
				channelName: channel.name,
				messageId: message.id,
				rootId,
				actorId: message.authorId,
				actorName: message.authorName,
				ts: message.ts,
				preview: activityPreview(message),
			};
			const rootAuthor =
				located.get(rootId)?.message.authorId ?? messages.find((m) => m.id === rootId)?.authorId;
			let kind: ActivityKind | undefined;
			if (mentionsUser(message, selfId)) kind = 'mention';
			else if (message.replyTo && (rootAuthor === selfId || input.isFollowed?.(channel.id, rootId)))
				kind = 'reply';
			if (kind) {
				const id = `${kind}:${message.id}`;
				items.push({ ...base, id, kind, read: isRead(id, message.ts) });
			}
		}
		for (const pin of input.pinsOf(channel.id)) {
			if (pin.pinnedBy === selfId || pin.message.deleted) continue;
			const id = `pin:${channel.id}:${pin.message.id}:${pin.pinnedAt}`;
			items.push({
				id,
				kind: 'pin',
				...(channel.kind === 'direct' ? { chat: true } : {}),
				channelId: channel.id,
				channelName: channel.name,
				messageId: pin.message.id,
				rootId: located.get(pin.message.id)?.rootId ?? pin.message.id,
				actorId: pin.pinnedBy,
				actorName: pin.pinnedByName,
				ts: pin.pinnedAt,
				preview: activityPreview(pin.message),
				read: isRead(id, pin.pinnedAt),
			});
		}
	}
	for (const reaction of input.reactions) {
		const target = located.get(reaction.messageId);
		if (!target || reaction.ts <= 0) continue;
		const { message, channel, rootId } = target;
		if (message.authorId !== selfId || message.deleted || reaction.userId === selfId) continue;
		const id = `reaction:${message.id}|${reaction.emoji}|${reaction.userId}`;
		items.push({
			id,
			kind: 'reaction',
			...(channel.kind === 'direct' ? { chat: true } : {}),
			channelId: channel.id,
			channelName: channel.name,
			messageId: message.id,
			rootId,
			actorId: reaction.userId,
			actorName: names.get(reaction.userId) ?? 'Someone',
			ts: reaction.ts,
			preview: activityPreview(message),
			emoji: reaction.emoji,
			read: isRead(id, reaction.ts),
		});
	}
	return items.sort((a, b) => b.ts - a.ts || a.id.localeCompare(b.id)).slice(0, MAX_ACTIVITY_ITEMS);
}

export type ActivityFilter = 'all' | 'unread' | 'mentions' | 'replies' | 'reactions';

export const ACTIVITY_FILTERS: readonly ActivityFilter[] = [
	'all',
	'unread',
	'mentions',
	'replies',
	'reactions',
];

/** The items a feed filter shows (Unread, Mentions, Replies and Reactions as in Teams). */
export function filterActivity(
	items: readonly ActivityItem[],
	filter: ActivityFilter,
): ActivityItem[] {
	if (filter === 'all') return [...items];
	if (filter === 'unread') return items.filter((item) => !item.read);
	const kind: ActivityKind =
		filter === 'mentions' ? 'mention' : filter === 'replies' ? 'reply' : 'reaction';
	return items.filter((item) => item.kind === kind);
}

/** The line a feed row shows, like Teams: who did what. */
export function describeActivity(item: ActivityItem): string {
	const where = item.chat ? item.channelName : `# ${item.channelName}`;
	switch (item.kind) {
		case 'mention':
			return `${item.actorName} mentioned you in ${where}`;
		case 'reply':
			return `${item.actorName} replied in ${where}`;
		case 'reaction':
			return `${item.actorName} reacted ${item.emoji ?? ''} to your message`.replace('  ', ' ');
		case 'pin':
			return `${item.actorName} pinned a message in ${where}`;
		case 'meeting':
			return `${item.channelName} starts soon`;
	}
}
