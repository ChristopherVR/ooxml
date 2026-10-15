// Pinned messages as a Yjs CRDT. One Y.Map per channel (`teams:pins:<channel>`) holds one row per
// pinned message id, so two people pinning the same message converge and a pin never rewrites the
// message itself. Any member can pin and unpin (like Teams); rows are re-validated when read back
// and a pin whose message is missing or deleted is simply not listed. DOM-free. New code.
import type * as Y from 'yjs';
import { isValidId } from '../collab/validation';
import { type Message, sanitizeChannelName } from './model';

export const MAX_PINNED_MESSAGES = 50;
const pinsMap = (channelId: string): string => `teams:pins:${channelId}`;

export interface PinnedMessage {
	message: Message;
	pinnedBy: string;
	pinnedByName: string;
	pinnedAt: number;
}

export interface PinStore {
	/** Pinned messages of a channel, most recently pinned first. */
	pinned: (channelId: string) => PinnedMessage[];
	/** Pins a live message of the channel. False when it is missing, deleted, pinned or the list is full. */
	pin: (channelId: string, messageId: string) => boolean;
	/** Returns whether a pin was removed. */
	unpin: (channelId: string, messageId: string) => boolean;
}

interface PinRow {
	by: string;
	byName: string;
	at: number;
}

function readRow(raw: unknown): PinRow | null {
	if (!raw || typeof raw !== 'object') return null;
	const r = raw as Record<string, unknown>;
	if (!isValidId(r.by)) return null;
	return {
		by: r.by,
		byName: sanitizeChannelName(r.byName) || 'Unknown',
		at: Number(r.at) || 0,
	};
}

export function createPinStore(
	doc: Y.Doc,
	user: { id: string; name: string },
	isChannel: (channelId: string) => boolean,
	messages: (channelId: string) => Message[],
): PinStore {
	const rows = (channelId: string): Y.Map<unknown> => doc.getMap<unknown>(pinsMap(channelId));
	const pinned: PinStore['pinned'] = (channelId) => {
		if (!isChannel(channelId)) return [];
		const byId = new Map(messages(channelId).map((m) => [m.id, m]));
		const out: PinnedMessage[] = [];
		rows(channelId).forEach((raw, messageId) => {
			const row = readRow(raw);
			const message = byId.get(messageId);
			if (!row || !message || message.deleted) return;
			out.push({ message, pinnedBy: row.by, pinnedByName: row.byName, pinnedAt: row.at });
		});
		return out
			.sort((a, b) => b.pinnedAt - a.pinnedAt || a.message.id.localeCompare(b.message.id))
			.slice(0, MAX_PINNED_MESSAGES);
	};
	return {
		pinned,
		pin(channelId, messageId) {
			if (!isChannel(channelId) || !isValidId(messageId)) return false;
			const target = messages(channelId).find((m) => m.id === messageId);
			if (!target || target.deleted) return false;
			const map = rows(channelId);
			if (map.has(messageId) || pinned(channelId).length >= MAX_PINNED_MESSAGES) return false;
			map.set(messageId, { by: user.id, byName: user.name, at: Date.now() } satisfies PinRow);
			return true;
		},
		unpin(channelId, messageId) {
			if (!isChannel(channelId)) return false;
			const map = rows(channelId);
			if (!map.has(messageId)) return false;
			map.delete(messageId);
			return true;
		},
	};
}
