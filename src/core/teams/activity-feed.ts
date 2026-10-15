// The activity feed of one client: derives the items (`activity.ts`), overlays personal read state
// (`activity-state.ts`) and decides which new items to announce (`notifications.ts`). The client
// calls `update` on every state change, then announces what `drain` returns. DOM-free.
import { type ActivityInput, type ActivityItem, deriveActivity } from './activity';
import { createActivityReadState } from './activity-state';
import { createNotificationSettings, shouldNotify } from './notifications';
import type { PresenceStatus } from './presence';
import type { StorageLike } from './store';

export interface ActivityFeed {
	/** Re-derive the feed; items that appeared since the last call are queued for `drain`. */
	update: (input: Omit<ActivityInput, 'isRead'>, status: PresenceStatus) => ActivityFeedView;
	/** The new, unread items that settings and presence allow to be announced. Clears the queue. */
	drain: () => ActivityItem[];
	markRead: (id: string, read?: boolean) => boolean;
	markAllRead: () => boolean;
	readonly settings: ReturnType<typeof createNotificationSettings>;
}

export interface ActivityFeedView {
	items: ActivityItem[];
	unread: number;
}

export function createActivityFeed(
	storage: StorageLike | undefined,
	keyBase: string,
): ActivityFeed {
	const read = createActivityReadState(storage, `${keyBase}:read`);
	const settings = createNotificationSettings(storage, `${keyBase}:notifications`);
	let items: ActivityItem[] = [];
	let seen: Set<string> | undefined;
	let queue: ActivityItem[] = [];
	return {
		settings,
		update(input, status) {
			items = deriveActivity({ ...input, isRead: read.isRead });
			if (!seen) seen = new Set(items.map((item) => item.id));
			else {
				const current = settings.get();
				for (const item of items) {
					if (seen.has(item.id)) continue;
					seen.add(item.id);
					if (shouldNotify(item, current, status)) queue.push(item);
				}
			}
			return { items, unread: items.filter((item) => !item.read).length };
		},
		drain() {
			const out = queue;
			queue = [];
			return out;
		},
		markRead(id, value = true) {
			const item = items.find((entry) => entry.id === id);
			return item ? read.mark(id, item.ts, value) : false;
		},
		markAllRead() {
			return read.markAll(Math.max(0, ...items.map((item) => item.ts)));
		},
	};
}
