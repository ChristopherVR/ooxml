// Channels and messages as a Yjs CRDT. Channels live in one Y.Map, messages in one Y.Map per
// channel, reactions in a flat Y.Map keyed `message|emoji|user` so two people reacting at the same
// time never overwrite each other. Everything a peer wrote is re-validated when read back, so a
// hostile or buggy client cannot inject markup or malformed rows. DOM-free. New code.
import * as Y from 'yjs';
import { createIdGenerator } from '../collab/identity';
import { isValidId } from '../collab/validation';
import { type Mention, type MentionTarget, resolveMentions, sanitizeMentions } from './mentions';
import { type PinnedMessage, createPinStore } from './pins';
import { tabConversationId } from './tab-conversation';
import type { ChannelTab } from './tabs';
import {
	type Attachment,
	type Channel,
	type ChannelKind,
	type Message,
	sanitizeAttachments,
	sanitizeChannelName,
	sanitizeChannelTopic,
	sanitizeEmoji,
	sanitizeMessageText,
} from './model';

const CHANNELS = 'teams:channels';
const REACTIONS = 'teams:reactions';
const messagesMap = (channelId: string): string => `teams:messages:${channelId}`;
const SEP = '|';

export interface ChatUser {
	id: string;
	name: string;
}

export interface ReactionEvent {
	messageId: string;
	emoji: string;
	userId: string;
	ts: number;
}

export interface ChatStore {
	channels: () => Channel[];
	createChannel: (input: {
		name: string;
		topic?: string;
		kind?: ChannelKind;
		id?: string;
	}) => Channel | null;
	archiveChannel: (channelId: string, archived?: boolean) => void;
	messages: (channelId: string) => Message[];
	/** Post a message; returns null when the text and attachments are both empty. */
	post: (
		channelId: string,
		input: {
			text: string;
			replyTo?: string;
			attachments?: Attachment[];
			/** Who may be mentioned; the `@name` runs of the text that match are stored as mentions. */
			mentions?: readonly MentionTarget[];
			/** Mark the row as an app notice (members added or removed). */
			system?: boolean;
			/** Announce a scheduled meeting with a card (the text is the fallback for old clients). */
			meetingId?: string;
		},
	) => Message | null;
	/** Idempotent root creation. Concurrent peers use one stable message id; existing rows survive. */
	ensureTabConversation: (tab: ChannelTab) => Message | null;
	/** Only the author can edit. Returns whether the edit applied. */
	edit: (
		channelId: string,
		messageId: string,
		text: string,
		mentions?: readonly MentionTarget[],
	) => boolean;
	/** Soft delete (the row stays so threads keep their shape). Only the author can delete. */
	remove: (channelId: string, messageId: string) => boolean;
	/** Pin a live message of the channel; any member may. Returns whether it applied. */
	pin: (channelId: string, messageId: string) => boolean;
	unpin: (channelId: string, messageId: string) => boolean;
	/** Pinned messages of a channel, most recently pinned first. */
	pinned: (channelId: string) => PinnedMessage[];
	/** Every live reaction with the time it was added (0 for rows written by older clients). */
	reactionEvents: () => ReactionEvent[];
	/** `channelId` routes to a private chat's own document; channels ignore it. */
	react: (messageId: string, emoji: string, channelId?: string) => void;
	unreact: (messageId: string, emoji: string, channelId?: string) => void;
	/** Fires after any channel, message or reaction change (local or remote), coalesced per tick. */
	observe: (listener: () => void) => () => void;
	destroy: () => void;
}

/** Message times from this client strictly increase, so two sends in one millisecond keep their order. */
let lastMessageTs = 0;
const nextMessageTs = (): number => (lastMessageTs = Math.max(Date.now(), lastMessageTs + 1));

const isChannelId = (id: unknown): id is string => isValidId(id, 64) && /^[\w-]+$/u.test(id);

function readChannel(id: string, raw: unknown): Channel | null {
	if (!(raw instanceof Y.Map)) return null;
	const name = sanitizeChannelName(raw.get('name'));
	if (!name || !isChannelId(id)) return null;
	return {
		id,
		name,
		topic: sanitizeChannelTopic(raw.get('topic')),
		kind: raw.get('kind') === 'direct' ? 'direct' : 'channel',
		createdAt: Number(raw.get('createdAt')) || 0,
		createdBy: String(raw.get('createdBy') ?? '').slice(0, 160),
		archived: raw.get('archived') === true,
	};
}

function readMessage(
	channelId: string,
	id: string,
	raw: unknown,
	reactions: Record<string, string[]>,
): Message | null {
	if (!(raw instanceof Y.Map) || !isValidId(id)) return null;
	const authorId = raw.get('authorId');
	if (!isValidId(authorId)) return null;
	const deleted = raw.get('deleted') === true;
	const replyTo = raw.get('replyTo');
	const editedAt = Number(raw.get('editedAt'));
	const message: Message = {
		id,
		channelId,
		authorId,
		authorName: sanitizeChannelName(raw.get('authorName')) || 'Unknown',
		text: deleted ? '' : sanitizeMessageText(raw.get('text')),
		ts: Number(raw.get('ts')) || 0,
		deleted,
		attachments: deleted ? [] : sanitizeAttachments(raw.get('attachments')),
		reactions,
	};
	if (isValidId(replyTo)) message.replyTo = replyTo;
	if (raw.get('system') === true) message.system = true;
	const meetingId = raw.get('meetingId');
	if (typeof meetingId === 'string' && /^[\w-]{1,64}$/u.test(meetingId) && !deleted)
		message.meetingId = meetingId;
	const tabId = raw.get('tabId');
	if (typeof tabId === 'string' && !message.replyTo && tabConversationId(tabId) === id)
		message.tabId = tabId;
	if (editedAt > 0) message.editedAt = editedAt;
	const mentions = sanitizeMentions(raw.get('mentions'), message.text);
	if (mentions.length) message.mentions = mentions;
	return message;
}

const slug = (name: string): string =>
	name
		.toLowerCase()
		.replace(/[^a-z0-9]+/gu, '-')
		.replace(/^-|-$/gu, '') || 'channel';

export function createChatStore(doc: Y.Doc, user: ChatUser): ChatStore {
	const nextId = createIdGenerator(String(doc.clientID), 'm');
	const channelMap = doc.getMap<Y.Map<unknown>>(CHANNELS);
	const reactionMap = doc.getMap<boolean | number>(REACTIONS);
	const channelMessages = (channelId: string): Y.Map<Y.Map<unknown>> =>
		doc.getMap<Y.Map<unknown>>(messagesMap(channelId));
	const listeners = new Set<() => void>();
	let scheduled = false;
	const onUpdate = (): void => {
		if (scheduled) return;
		scheduled = true;
		queueMicrotask(() => {
			scheduled = false;
			for (const l of [...listeners]) l();
		});
	};
	doc.on('update', onUpdate);

	const reactionsFor = (messageId: string): Record<string, string[]> => {
		const out: Record<string, string[]> = {};
		const prefix = messageId + SEP;
		reactionMap.forEach((on, key) => {
			if (!on || !key.startsWith(prefix)) return;
			const [, emoji, userId] = key.split(SEP);
			if (!emoji || !userId || !sanitizeEmoji(emoji)) return;
			(out[emoji] ??= []).push(userId);
		});
		return out;
	};

	const messagesOf = (channelId: string): Message[] => {
		const out: Message[] = [];
		if (!isChannelId(channelId)) return out;
		channelMessages(channelId).forEach((raw, id) => {
			const m = readMessage(channelId, id, raw, {});
			if (m) out.push(m);
		});
		return out;
	};
	const pins = createPinStore(doc, user, (id) => isChannelId(id) && channelMap.has(id), messagesOf);

	return {
		channels() {
			const out: Channel[] = [];
			channelMap.forEach((raw, id) => {
				const c = readChannel(id, raw);
				if (c) out.push(c);
			});
			return out.sort((a, b) => a.createdAt - b.createdAt || a.name.localeCompare(b.name));
		},
		createChannel({ name, topic = '', kind = 'channel', id }) {
			const clean = sanitizeChannelName(name);
			if (!clean) return null;
			const channelId = id && isChannelId(id) ? id : `${slug(clean)}-${nextId('c').slice(-6)}`;
			if (channelMap.has(channelId)) return readChannel(channelId, channelMap.get(channelId));
			doc.transact(() => {
				const m = new Y.Map<unknown>();
				m.set('name', clean);
				m.set('topic', sanitizeChannelTopic(topic));
				m.set('kind', kind);
				m.set('createdAt', Date.now());
				m.set('createdBy', user.id);
				m.set('archived', false);
				channelMap.set(channelId, m);
			});
			return readChannel(channelId, channelMap.get(channelId));
		},
		archiveChannel(channelId, archived = true) {
			channelMap.get(channelId)?.set('archived', archived);
		},
		messages(channelId) {
			if (!isChannelId(channelId)) return [];
			const out: Message[] = [];
			channelMessages(channelId).forEach((raw, id) => {
				const m = readMessage(channelId, id, raw, reactionsFor(id));
				if (m) out.push(m);
			});
			return out.sort((a, b) => a.ts - b.ts || a.id.localeCompare(b.id));
		},
		post(channelId, { text, replyTo, attachments, mentions, system, meetingId }) {
			if (!isChannelId(channelId) || !channelMap.has(channelId)) return null;
			const clean = sanitizeMessageText(text);
			const files = sanitizeAttachments(attachments);
			if (!clean && files.length === 0) return null;
			const id = nextId('m');
			doc.transact(() => {
				const m = new Y.Map<unknown>();
				m.set('authorId', user.id);
				m.set('authorName', user.name);
				m.set('text', clean);
				m.set('ts', nextMessageTs());
				m.set('deleted', false);
				if (system) m.set('system', true);
				if (meetingId && /^[\w-]{1,64}$/u.test(meetingId)) m.set('meetingId', meetingId);
				if (replyTo && isValidId(replyTo)) m.set('replyTo', replyTo);
				if (files.length) m.set('attachments', files);
				const found = resolveMentions(clean, mentions ?? []);
				if (found.length) m.set('mentions', found);
				channelMessages(channelId).set(id, m);
			});
			return readMessage(channelId, id, channelMessages(channelId).get(id), {});
		},
		ensureTabConversation(tab) {
			const id = tabConversationId(tab.id);
			const name = sanitizeChannelName(tab.name);
			if (!id || !name || !isChannelId(tab.channelId) || !channelMap.has(tab.channelId))
				return null;
			const rows = channelMessages(tab.channelId);
			if (!rows.has(id))
				doc.transact(() => {
					const row = new Y.Map<unknown>();
					row.set('authorId', user.id);
					row.set('authorName', user.name);
					row.set('text', `Discuss the ${name} tab here.`);
					row.set('ts', Date.now());
					row.set('deleted', false);
					row.set('tabId', tab.id);
					rows.set(id, row);
				});
			const message = readMessage(tab.channelId, id, rows.get(id), reactionsFor(id));
			return message?.tabId === tab.id ? message : null;
		},
		edit(channelId, messageId, text, mentions) {
			if (!isChannelId(channelId)) return false;
			const m = channelMessages(channelId).get(messageId);
			const clean = sanitizeMessageText(text);
			if (!m || m.get('authorId') !== user.id || m.get('deleted') === true || !clean) return false;
			doc.transact(() => {
				// Mentions already in the message stay addressable, so editing around one keeps it.
				const kept = sanitizeMentions(m.get('mentions'), sanitizeMessageText(m.get('text')));
				const found: Mention[] = resolveMentions(clean, [...(mentions ?? []), ...kept]);
				m.set('text', clean);
				if (found.length) m.set('mentions', found);
				else m.delete('mentions');
				m.set('editedAt', Date.now());
			});
			return true;
		},
		remove(channelId, messageId) {
			if (!isChannelId(channelId)) return false;
			const m = channelMessages(channelId).get(messageId);
			if (!m || m.get('authorId') !== user.id) return false;
			m.set('deleted', true);
			return true;
		},
		pin: (channelId, messageId) =>
			channelMap.get(channelId)?.get('archived') !== true && pins.pin(channelId, messageId),
		unpin: (channelId, messageId) => pins.unpin(channelId, messageId),
		pinned: (channelId) => pins.pinned(channelId),
		reactionEvents() {
			const out: ReactionEvent[] = [];
			reactionMap.forEach((on, key) => {
				const [messageId, emoji, userId] = key.split(SEP);
				if (!on || !messageId || !emoji || !userId || !sanitizeEmoji(emoji)) return;
				out.push({ messageId, emoji, userId, ts: typeof on === 'number' ? on : 0 });
			});
			return out;
		},
		react(messageId, emoji) {
			const e = sanitizeEmoji(emoji);
			if (e && isValidId(messageId)) reactionMap.set([messageId, e, user.id].join(SEP), Date.now());
		},
		unreact(messageId, emoji) {
			const e = sanitizeEmoji(emoji);
			if (e) reactionMap.delete([messageId, e, user.id].join(SEP));
		},
		observe(listener) {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
		destroy() {
			doc.off('update', onUpdate);
			listeners.clear();
		},
	};
}
