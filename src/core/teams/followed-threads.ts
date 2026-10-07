import { isValidId } from '../collab/validation.js';
import type { Channel, Message } from './model.js';
import type { StorageLike } from './store.js';
import { channelThreads } from './threads.js';

export interface FollowedThread {
	channelId: string;
	channelName: string;
	root: Message;
	replyCount: number;
	updatedAt: number;
}

/** Personal references stay outside Yjs. Storage failures retain the in-memory preference. */
export function createThreadFollows(storage: StorageLike | undefined, key: string) {
	const follows = new Map<string, { channelId: string; messageId: string }>();
	const reference = (channel: string, message: string) => JSON.stringify([channel, message]);
	const matching = (
		channel: string,
		message: string,
		rootOf?: (id: string) => string | undefined,
	) =>
		[...follows.entries()]
			.filter(
				([, row]) =>
					row.channelId === channel &&
					(row.messageId === message || rootOf?.(row.messageId) === message),
			)
			.map(([id]) => id);
	try {
		const raw = storage?.getItem(key) ?? '[]';
		const rows: unknown = raw.length <= 131_072 ? JSON.parse(raw) : [];
		if (Array.isArray(rows))
			for (const row of rows.slice(0, 250)) {
				if (
					row &&
					typeof row === 'object' &&
					isValidId(row.channelId, 64) &&
					isValidId(row.messageId)
				)
					follows.set(reference(row.channelId, row.messageId), {
						channelId: row.channelId,
						messageId: row.messageId,
					});
			}
	} catch {
		/* Corrupt or blocked local preferences do not affect the shared document. */
	}
	return {
		has(channel: string, message: string, rootOf?: (id: string) => string | undefined): boolean {
			return matching(channel, message, rootOf).length > 0;
		},
		set(
			channel: string,
			message: string,
			followed: boolean,
			rootOf?: (id: string) => string | undefined,
		): boolean {
			if (!isValidId(channel, 64) || !isValidId(message)) return false;
			const id = reference(channel, message);
			const matches = matching(channel, message, rootOf);
			if (followed === matches.length > 0 || (followed && follows.size >= 250)) return false;
			if (followed) follows.set(id, { channelId: channel, messageId: message });
			else for (const match of matches) follows.delete(match);
			try {
				storage?.setItem(key, JSON.stringify([...follows.values()]));
			} catch {
				/* In-memory only. */
			}
			return true;
		},
		view(channels: readonly Channel[], messagesOf: (id: string) => Message[]): FollowedThread[] {
			const visible = new Map(
				channels.filter((channel) => !channel.archived).map((channel) => [channel.id, channel]),
			);
			const indexes = new Map<string, ReturnType<typeof channelThreads>>();
			const result: FollowedThread[] = [];
			const seen = new Set<string>();
			for (const { channelId, messageId } of follows.values()) {
				const channel = visible.get(channelId);
				if (!channel) continue;
				let index = indexes.get(channelId);
				if (!index) {
					index = channelThreads(messagesOf(channelId));
					indexes.set(channelId, index);
				}
				const thread = index.thread(messageId);
				if (!thread) continue;
				const id = reference(channelId, thread.root.id);
				if (seen.has(id)) continue;
				seen.add(id);
				let updatedAt = Math.max(thread.root.ts, thread.root.editedAt ?? 0);
				for (const reply of thread.replies)
					if (!reply.deleted) updatedAt = Math.max(updatedAt, reply.ts, reply.editedAt ?? 0);
				result.push({
					channelId,
					channelName: channel.name,
					root: thread.root,
					replyCount: index.replyCounts[thread.root.id] ?? 0,
					updatedAt,
				});
			}
			return result.sort((a, b) => b.updatedAt - a.updatedAt || a.root.id.localeCompare(b.root.id));
		},
	};
}
