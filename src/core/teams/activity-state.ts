// Personal read state of the activity feed. It stays outside Yjs, like followed threads and read
// markers: what you have read is yours. The model is a watermark (`readAt`: everything up to it is
// read, "Mark all as read" raises it) plus explicit overrides for single items. Storage failures
// keep the in-memory state. A first run starts at "now", so history that predates this device does
// not show up as a wall of unread items. DOM-free.
import { isValidId } from '../collab/validation';
import type { StorageLike } from './store';

const MAX_OVERRIDES = 500;
const MAX_RAW_CHARS = 131_072;

export interface ActivityReadState {
	isRead: (id: string, ts: number) => boolean;
	/** Mark one item read (or unread). Returns whether anything changed. */
	mark: (id: string, ts: number, read: boolean) => boolean;
	/** Everything up to `upTo` is read; overrides below it are dropped. */
	markAll: (upTo: number) => boolean;
	/** The watermark: items at or before it are read unless overridden. */
	readAt: () => number;
}

interface Persisted {
	readAt: number;
	read: string[];
	unread: string[];
}

export function createActivityReadState(
	storage: StorageLike | undefined,
	key: string,
	now: () => number = Date.now,
): ActivityReadState {
	let readAt = now();
	const read = new Set<string>();
	const unread = new Set<string>();
	let restored = false;
	try {
		const raw = storage?.getItem(key);
		const parsed: unknown = raw && raw.length <= MAX_RAW_CHARS ? JSON.parse(raw) : null;
		if (parsed && typeof parsed === 'object') {
			const row = parsed as Partial<Persisted>;
			if (typeof row.readAt === 'number' && Number.isFinite(row.readAt) && row.readAt >= 0) {
				readAt = row.readAt;
				restored = true;
			}
			for (const [list, target] of [
				[row.read, read],
				[row.unread, unread],
			] as const)
				if (Array.isArray(list))
					for (const id of list.slice(0, MAX_OVERRIDES)) if (isValidId(id, 400)) target.add(id);
		}
	} catch {
		/* Corrupt or blocked preferences fall back to a fresh start. */
	}
	const persist = (): void => {
		try {
			storage?.setItem(
				key,
				JSON.stringify({ readAt, read: [...read], unread: [...unread] } satisfies Persisted),
			);
		} catch {
			/* In-memory only. */
		}
	};
	if (!restored) persist();
	const trim = (set: Set<string>): void => {
		while (set.size > MAX_OVERRIDES) set.delete(set.values().next().value as string);
	};
	const isRead = (id: string, ts: number): boolean =>
		!unread.has(id) && (read.has(id) || ts <= readAt);
	return {
		readAt: () => readAt,
		isRead,
		mark(id, ts, value) {
			if (!isValidId(id, 400)) return false;
			if (isRead(id, ts) === value) return false;
			if (value) {
				unread.delete(id);
				if (ts > readAt) read.add(id);
			} else {
				read.delete(id);
				unread.add(id);
			}
			trim(read);
			trim(unread);
			persist();
			return true;
		},
		markAll(upTo) {
			if (!Number.isFinite(upTo) || (upTo <= readAt && read.size === 0 && unread.size === 0))
				return false;
			readAt = Math.max(readAt, upTo);
			read.clear();
			unread.clear();
			persist();
			return true;
		},
	};
}
