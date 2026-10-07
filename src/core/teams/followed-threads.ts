import { isValidId } from '../collab/validation';
import type { Channel, Message } from './model';
import type { StorageLike } from './store';
import { channelThreads } from './threads';

export interface FollowedThread {
	channelId: string;
	channelName: string;
	root: Message;
	replyCount: number;
	updatedAt: number;
	unread: number;
}

export interface ThreadFollowSettings {
	started: boolean;
	replied: boolean;
}

/** Personal references stay outside Yjs. Storage failures retain the in-memory preference. */
export function createThreadFollows(storage: StorageLike | undefined, key: string) {
	const settings: ThreadFollowSettings = { started: true, replied: true };
	try {
		const raw = storage?.getItem(`${key}:settings`) ?? '{}';
		const restored: unknown = raw.length <= 1024 ? JSON.parse(raw) : {};
		if (restored && typeof restored === 'object')
			for (const name of ['started', 'replied'] as const) {
				const value = (restored as Record<string, unknown>)[name];
				if (typeof value === 'boolean') settings[name] = value;
			}
	} catch {
		/* Defaults remain usable without storage. */
	}
	const follows = new Map<
		string,
		{ channelId: string; messageId: string; readAt: number; forcedUnread: boolean }
	>();
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
						readAt:
							typeof row.readAt === 'number' && Number.isFinite(row.readAt) && row.readAt >= 0
								? row.readAt
								: 0,
						forcedUnread: row.forcedUnread === true,
					});
			}
	} catch {
		/* Corrupt or blocked local preferences do not affect the shared document. */
	}
	const persist = () => {
		try {
			storage?.setItem(key, JSON.stringify([...follows.values()]));
		} catch {
			/* In-memory only. */
		}
	};
	return {
		settings: (): ThreadFollowSettings => ({ ...settings }),
		configure(input: Partial<ThreadFollowSettings>): boolean {
			let changed = false;
			for (const name of ['started', 'replied'] as const)
				if (typeof input[name] === 'boolean' && input[name] !== settings[name]) {
					settings[name] = input[name];
					changed = true;
				}
			if (changed)
				try {
					storage?.setItem(`${key}:settings`, JSON.stringify(settings));
				} catch {
					/* In-memory only. */
				}
			return changed;
		},
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
			if (followed)
				follows.set(id, { channelId: channel, messageId: message, readAt: 0, forcedUnread: false });
			else for (const match of matches) follows.delete(match);
			persist();
			return true;
		},
		mark(
			channel: string,
			message: string,
			readAt: number,
			forcedUnread: boolean,
			rootOf?: (id: string) => string | undefined,
		): boolean {
			let changed = false;
			for (const id of matching(channel, message, rootOf)) {
				const row = follows.get(id)!;
				const nextReadAt = forcedUnread ? row.readAt : readAt;
				if (row.readAt === nextReadAt && row.forcedUnread === forcedUnread) continue;
				row.readAt = nextReadAt;
				row.forcedUnread = forcedUnread;
				changed = true;
			}
			if (changed) persist();
			return changed;
		},
		view(
			channels: readonly Channel[],
			messagesOf: (id: string) => Message[],
			selfId = '',
		): FollowedThread[] {
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
				const preferences = matching(
					channelId,
					thread.root.id,
					(id) => index!.thread(id)?.root.id,
				).map((id) => follows.get(id)!);
				const readAt = Math.min(...preferences.map((row) => row.readAt));
				const unread = [thread.root, ...thread.replies].filter(
					(message) => !message.deleted && message.authorId !== selfId && message.ts > readAt,
				).length;
				result.push({
					channelId,
					channelName: channel.name,
					root: thread.root,
					replyCount: index.replyCounts[thread.root.id] ?? 0,
					updatedAt,
					unread: preferences.some((row) => row.forcedUnread) ? Math.max(1, unread) : unread,
				});
			}
			return result.sort((a, b) => b.updatedAt - a.updatedAt || a.root.id.localeCompare(b.root.id));
		},
	};
}
