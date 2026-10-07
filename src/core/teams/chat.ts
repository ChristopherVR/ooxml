// Channels and messages as a Yjs CRDT. Channels live in one Y.Map, messages in one Y.Map per
// channel, reactions in a flat Y.Map keyed `message|emoji|user` so two people reacting at the same
// time never overwrite each other. Everything a peer wrote is re-validated when read back, so a
// hostile or buggy client cannot inject markup or malformed rows. DOM-free. New code.
import * as Y from 'yjs';
import { createIdGenerator } from '../collab/identity';
import { isValidId } from '../collab/validation';
import { tabConversationId } from './tab-conversation';
import type { ChannelTab } from './tabs';
import {
	type Attachment,
	type Channel,
	type ChannelKind,
	type Message,
	sanitizeAttachments,
	sanitizeChannelName,
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
		input: { text: string; replyTo?: string; attachments?: Attachment[] },
	) => Message | null;
	/** Idempotent root creation. Concurrent peers use one stable message id; existing rows survive. */
	ensureTabConversation: (tab: ChannelTab) => Message | null;
	/** Only the author can edit. Returns whether the edit applied. */
	edit: (channelId: string, messageId: string, text: string) => boolean;
	/** Soft delete (the row stays so threads keep their shape). Only the author can delete. */
	remove: (channelId: string, messageId: string) => boolean;
	react: (messageId: string, emoji: string) => void;
	unreact: (messageId: string, emoji: string) => void;
	/** Fires after any channel, message or reaction change (local or remote), coalesced per tick. */
	observe: (listener: () => void) => () => void;
	destroy: () => void;
}

const isChannelId = (id: unknown): id is string => isValidId(id, 64) && /^[\w-]+$/u.test(id);

function readChannel(id: string, raw: unknown): Channel | null {
	if (!(raw instanceof Y.Map)) return null;
	const name = sanitizeChannelName(raw.get('name'));
	if (!name || !isChannelId(id)) return null;
	return {
		id,
		name,
		topic: sanitizeChannelName(raw.get('topic')),
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
	const tabId = raw.get('tabId');
	if (typeof tabId === 'string' && !message.replyTo && tabConversationId(tabId) === id)
		message.tabId = tabId;
	if (editedAt > 0) message.editedAt = editedAt;
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
	const reactionMap = doc.getMap<boolean>(REACTIONS);
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
				m.set('topic', sanitizeChannelName(topic));
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
		post(channelId, { text, replyTo, attachments }) {
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
				m.set('ts', Date.now());
				m.set('deleted', false);
				if (replyTo && isValidId(replyTo)) m.set('replyTo', replyTo);
				if (files.length) m.set('attachments', files);
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
		edit(channelId, messageId, text) {
			if (!isChannelId(channelId)) return false;
			const m = channelMessages(channelId).get(messageId);
			const clean = sanitizeMessageText(text);
			if (!m || m.get('authorId') !== user.id || m.get('deleted') === true || !clean) return false;
			doc.transact(() => {
				m.set('text', clean);
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
		react(messageId, emoji) {
			const e = sanitizeEmoji(emoji);
			if (e && isValidId(messageId)) reactionMap.set([messageId, e, user.id].join(SEP), true);
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
