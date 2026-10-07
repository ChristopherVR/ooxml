import type { Message } from './model';

export interface MessageThread {
	root: Message;
	replies: Message[];
}

/** Resolve reply chains once. Missing parents and cycles retain a visible, deterministic root. */
export function channelThreads(messages: readonly Message[]): {
	posts: Message[];
	replyCounts: Record<string, number>;
	thread(messageId: string): MessageThread | null;
} {
	const byId = new Map(messages.map((message) => [message.id, message]));
	const roots = new Map<string, string>();
	for (const message of messages) {
		const path: string[] = [];
		const positions = new Map<string, number>();
		let current = message;
		let root: string;
		while (true) {
			const known = roots.get(current.id);
			if (known !== undefined) {
				root = known;
				break;
			}
			const position = positions.get(current.id);
			if (position !== undefined) {
				root = path.slice(position).sort()[0]!;
				break;
			}
			positions.set(current.id, path.length);
			path.push(current.id);
			const parent = current.replyTo ? byId.get(current.replyTo) : undefined;
			if (!parent) {
				root = current.id;
				break;
			}
			current = parent;
		}
		for (const id of path) roots.set(id, root);
	}
	const replies = new Map<string, Message[]>();
	const posts: Message[] = [];
	const replyCounts: Record<string, number> = Object.create(null) as Record<string, number>;
	for (const message of messages) {
		const root = roots.get(message.id)!;
		if (root === message.id) posts.push(message);
		else {
			const rows = replies.get(root) ?? [];
			rows.push(message);
			replies.set(root, rows);
			replyCounts[root] = (replyCounts[root] ?? 0) + (message.deleted ? 0 : 1);
		}
	}
	return {
		posts,
		replyCounts,
		thread(id) {
			const rootId = roots.get(id);
			const root = rootId ? byId.get(rootId) : undefined;
			return root ? { root, replies: replies.get(root.id) ?? [] } : null;
		},
	};
}
