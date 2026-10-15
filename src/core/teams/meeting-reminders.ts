// Meeting reminders: five minutes before an occurrence starts, once per device. What was already
// announced is remembered locally (a reload must not repeat a toast). Turning a reminder into a
// notification goes through the same gate as every activity item (`shouldNotify`: Off and Do not
// disturb silence it), so this module only decides *when*. DOM-free. New code.
import type { ActivityItem } from './activity';
import { REMINDER_LEAD_MS, remindersDue } from './calendar';
import { type Meeting, meetingVisibleTo } from './meetings';
import type { StorageLike } from './store';

const MAX_KEPT = 200;
const KEEP_MS = 24 * 3_600_000;

export interface MeetingReminders {
	/**
	 * Reminders that are due now and were not announced before. `allowed` says whether the person's
	 * settings and status let an item through; a reminder it refuses (Do not disturb) is not marked
	 * announced, so it still fires if the status changes before the meeting starts.
	 */
	due: (
		meetings: readonly Meeting[],
		selfId: string,
		now: number,
		allowed: (item: ActivityItem) => boolean,
	) => ActivityItem[];
}

export function createMeetingReminders(
	storage: StorageLike | undefined,
	key: string,
): MeetingReminders {
	/** occurrence key -> its start, so old entries can be pruned. */
	let seen: Record<string, number> = {};
	try {
		const raw = storage?.getItem(key);
		const parsed: unknown = raw && raw.length < 65_536 ? JSON.parse(raw) : null;
		if (parsed && typeof parsed === 'object')
			for (const [k, v] of Object.entries(parsed as Record<string, unknown>))
				if (typeof v === 'number' && Number.isFinite(v)) seen[k] = v;
	} catch {
		seen = {};
	}
	return {
		due(meetings, selfId, now, allowed) {
			const visible = meetings.filter((m) => meetingVisibleTo(m, selfId));
			const items: ActivityItem[] = [];
			for (const { meeting, occurrence } of remindersDue(visible, now, REMINDER_LEAD_MS)) {
				if (seen[occurrence.key] !== undefined) continue;
				const minutes = Math.max(1, Math.round((occurrence.start - now) / 60_000));
				const item: ActivityItem = {
					id: `meeting:${occurrence.key}`,
					kind: 'meeting',
					channelId: meeting.chatId ?? meeting.channelId ?? '',
					channelName: meeting.title,
					messageId: meeting.id,
					rootId: meeting.id,
					actorId: meeting.organizerId,
					actorName: meeting.organizerName,
					ts: now,
					preview: minutes === 1 ? 'Starts in 1 minute' : `Starts in ${minutes} minutes`,
					read: false,
				};
				if (!allowed(item)) continue;
				seen[occurrence.key] = occurrence.start;
				items.push(item);
			}
			if (items.length) {
				const kept = Object.entries(seen)
					.filter(([, start]) => start > now - KEEP_MS)
					.sort((a, b) => b[1] - a[1])
					.slice(0, MAX_KEPT);
				seen = Object.fromEntries(kept);
				try {
					storage?.setItem(key, JSON.stringify(seen));
				} catch {
					/* In-memory only. */
				}
			}
			return items;
		},
	};
}
