import * as Y from 'yjs';
import { createIdGenerator } from '../collab/identity.js';
import { isValidId } from '../collab/validation.js';
import { contentUrl } from './content.js';
import { sanitizeAttachment, sanitizeChannelName, type Attachment } from './model.js';
import type { ChatUser } from './chat.js';

export type TabContent =
	| { type: 'file'; attachment: Attachment }
	| { type: 'website'; url: string };
export interface ChannelTab {
	id: string;
	channelId: string;
	name: string;
	content: TabContent;
	createdBy: string;
	createdAt: number;
}
export interface TabStore {
	tabs(channelId: string): ChannelTab[];
	add(channelId: string, name: string, content: TabContent): ChannelTab | null;
	/** Advisory owner checks, not a replacement for server authorization. */
	rename(id: string, name: string): boolean;
	remove(id: string): boolean;
}

function readContent(raw: unknown): TabContent | null {
	if (!raw || typeof raw !== 'object') return null;
	const value = raw as Record<string, unknown>;
	if (value.type === 'file') {
		const attachment = sanitizeAttachment(value.attachment);
		if (attachment?.url && contentUrl(attachment.url, 'https://workspace.invalid'))
			return { type: 'file', attachment };
	}
	if (value.type === 'website' && typeof value.url === 'string' && value.url.length <= 2048) {
		if (contentUrl(value.url, 'https://workspace.invalid'))
			return { type: 'website', url: value.url };
	}
	return null;
}

const validId = (id: unknown): id is string => typeof id === 'string' && /^[\w-]{1,160}$/u.test(id);

/** One shared CRDT map; row fields merge independently (for example name and content). */
export function createTabStore(
	doc: Y.Doc,
	user: ChatUser,
	channelExists: (id: string) => boolean,
): TabStore {
	const rows = doc.getMap<Y.Map<unknown>>('teams:tabs');
	const nextId = createIdGenerator(String(doc.clientID), 'tab');
	function read(id: string, row: unknown): ChannelTab | null {
		if (!validId(id) || !(row instanceof Y.Map)) return null;
		const channelId = row.get('channelId');
		const createdBy = row.get('createdBy');
		const createdAt = row.get('createdAt');
		const name = sanitizeChannelName(row.get('name'));
		const content = readContent(row.get('content'));
		if (
			!validId(channelId) ||
			!channelExists(channelId) ||
			!isValidId(createdBy) ||
			!name ||
			!content ||
			typeof createdAt !== 'number' ||
			!Number.isFinite(createdAt) ||
			createdAt < 0
		)
			return null;
		return { id, channelId, name, content, createdBy, createdAt };
	}
	return {
		tabs(channelId) {
			if (!channelExists(channelId)) return [];
			const tabs: ChannelTab[] = [];
			rows.forEach((row, id) => {
				const tab = read(id, row);
				if (tab?.channelId === channelId) tabs.push(tab);
			});
			return tabs.sort((a, b) => a.createdAt - b.createdAt || a.id.localeCompare(b.id));
		},
		add(channelId, name, content) {
			const cleanName = sanitizeChannelName(name);
			const cleanContent = readContent(content);
			if (!channelExists(channelId) || !cleanName || !cleanContent) return null;
			const id = nextId('tab');
			doc.transact(() => {
				const row = new Y.Map<unknown>();
				row.set('channelId', channelId);
				row.set('name', cleanName);
				row.set('content', cleanContent);
				row.set('createdBy', user.id);
				row.set('createdAt', Date.now());
				rows.set(id, row);
			});
			return read(id, rows.get(id));
		},
		rename(id, name) {
			const tab = read(id, rows.get(id));
			const clean = sanitizeChannelName(name);
			if (!tab || tab.createdBy !== user.id || !clean) return false;
			rows.get(id)!.set('name', clean);
			return true;
		},
		remove(id) {
			const tab = read(id, rows.get(id));
			if (!tab || tab.createdBy !== user.id) return false;
			rows.delete(id);
			return true;
		},
	};
}
