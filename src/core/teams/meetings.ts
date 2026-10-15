// Scheduled meetings in the workspace Yjs document. A meeting stores its start and end as UTC
// instants plus the organiser's IANA zone, which only matters for repeats (a daily 09:00 stays 09:00
// on the organiser's clock across daylight saving); every viewer renders in their own zone. Rows
// are re-validated when read back. Any member can read the shared document, so the attendee list
// scopes who sees a meeting in their calendar, not who could read the row. DOM-free. New code.
import * as Y from 'yjs';
import { createIdGenerator } from '../collab/identity';
import { type ChatMember, isChatUserId, sanitizeMembers } from './chats';
import { sanitizeChannelName, sanitizeMessageText } from './model';
import { isValidZone, localZone } from './time-zone';

export type MeetingRepeat = 'none' | 'daily' | 'weekly';
export const MEETING_REPEATS: readonly MeetingRepeat[] = ['none', 'daily', 'weekly'];

export interface Meeting {
	id: string;
	title: string;
	description: string;
	organizerId: string;
	organizerName: string;
	/** The organiser's IANA zone: repeats keep their wall-clock time in it. */
	zone: string;
	/** UTC milliseconds. */
	start: number;
	end: number;
	repeat: MeetingRepeat;
	attendees: ChatMember[];
	/** A channel meeting: the channel it was scheduled in (its chat is the channel). */
	channelId?: string;
	/** The group chat of a meeting with attendees. */
	chatId?: string;
	createdAt: number;
	cancelled: boolean;
}

export interface MeetingInput {
	title: string;
	description?: string;
	start: number;
	end: number;
	repeat?: MeetingRepeat;
	attendees?: readonly ChatMember[];
	channelId?: string;
	chatId?: string;
	/** The organiser's zone; defaults to the device's. */
	zone?: string;
}

export const MAX_MEETING_TITLE = 120;
export const MAX_MEETING_DESCRIPTION = 2000;
export const MAX_MEETING_HOURS = 24;
const MEETINGS = 'teams:meetings';
const SAFE_ID = /^[\w-]{1,64}$/u;

export interface MeetingStore {
	list: () => Meeting[];
	get: (id: string) => Meeting | null;
	/** Null when the title is empty or the times are not a positive span of at most 24 hours. */
	create: (input: MeetingInput) => Meeting | null;
	/** Organiser only. Returns the updated meeting. */
	update: (id: string, patch: Partial<MeetingInput>) => Meeting | null;
	/** Organiser only. The row stays so cards and reminders show it as cancelled. */
	cancel: (id: string) => boolean;
}

const finiteTime = (v: unknown): number | null =>
	typeof v === 'number' && Number.isFinite(v) && v > 0 && v < 8.64e15 ? Math.floor(v) : null;

function readMeeting(id: string, raw: unknown): Meeting | null {
	if (!(raw instanceof Y.Map) || !SAFE_ID.test(id)) return null;
	const start = finiteTime(raw.get('start'));
	const end = finiteTime(raw.get('end'));
	const title = sanitizeChannelName(raw.get('title')).slice(0, MAX_MEETING_TITLE);
	const organizerId = raw.get('organizerId');
	if (start === null || end === null || end <= start || !title || !isChatUserId(organizerId))
		return null;
	if (end - start > MAX_MEETING_HOURS * 3_600_000) return null;
	const zone = raw.get('zone');
	const repeat = raw.get('repeat');
	const channelId = raw.get('channelId');
	const chatId = raw.get('chatId');
	return {
		id,
		title,
		description: sanitizeMessageText(raw.get('description')).slice(0, MAX_MEETING_DESCRIPTION),
		organizerId,
		organizerName: sanitizeChannelName(raw.get('organizerName')) || 'Unknown',
		zone: isValidZone(zone) ? zone : 'UTC',
		start,
		end,
		repeat: MEETING_REPEATS.includes(repeat as MeetingRepeat) ? (repeat as MeetingRepeat) : 'none',
		attendees: sanitizeMembers(raw.get('attendees')),
		createdAt: Number(raw.get('createdAt')) || 0,
		cancelled: raw.get('cancelled') === true,
		...(typeof channelId === 'string' && SAFE_ID.test(channelId) ? { channelId } : {}),
		...(typeof chatId === 'string' && SAFE_ID.test(chatId) ? { chatId } : {}),
	};
}

/** Whether a meeting belongs in this person's calendar. Channel meetings show for everyone. */
export const meetingVisibleTo = (meeting: Meeting, userId: string): boolean =>
	!meeting.cancelled &&
	(meeting.channelId !== undefined ||
		meeting.organizerId === userId ||
		meeting.attendees.some((a) => a.id === userId));

export function createMeetingStore(
	doc: Y.Doc,
	user: ChatMember,
	now: () => number = Date.now,
): MeetingStore {
	const map = doc.getMap<Y.Map<unknown>>(MEETINGS);
	const nextId = createIdGenerator(String(doc.clientID), 't');
	const get = (id: string): Meeting | null => readMeeting(id, map.get(id));
	const apply = (row: Y.Map<unknown>, input: Partial<MeetingInput>): void => {
		if (input.title !== undefined) row.set('title', sanitizeChannelName(input.title));
		if (input.description !== undefined)
			row.set('description', sanitizeMessageText(input.description));
		if (input.start !== undefined) row.set('start', input.start);
		if (input.end !== undefined) row.set('end', input.end);
		if (input.repeat !== undefined && MEETING_REPEATS.includes(input.repeat))
			row.set('repeat', input.repeat);
		if (input.attendees !== undefined)
			row.set(
				'attendees',
				sanitizeMembers(input.attendees).filter((a) => a.id !== user.id),
			);
		if (input.channelId !== undefined) {
			if (SAFE_ID.test(input.channelId)) row.set('channelId', input.channelId);
			else row.delete('channelId');
		}
		if (input.chatId && SAFE_ID.test(input.chatId)) row.set('chatId', input.chatId);
		if (input.zone !== undefined && isValidZone(input.zone)) row.set('zone', input.zone);
	};
	return {
		list() {
			const out: Meeting[] = [];
			map.forEach((raw, id) => {
				const m = readMeeting(id, raw);
				if (m) out.push(m);
			});
			return out.sort((a, b) => a.start - b.start || a.id.localeCompare(b.id));
		},
		get,
		create(input) {
			if (!isChatUserId(user.id)) return null;
			const start = finiteTime(input.start);
			const end = finiteTime(input.end);
			if (!sanitizeChannelName(input.title) || start === null || end === null) return null;
			if (end <= start || end - start > MAX_MEETING_HOURS * 3_600_000) return null;
			const id = nextId('t').replace(/[^\w-]/gu, '');
			doc.transact(() => {
				const row = new Y.Map<unknown>();
				row.set('organizerId', user.id);
				row.set('organizerName', user.name);
				row.set('createdAt', now());
				row.set('cancelled', false);
				row.set('zone', isValidZone(input.zone) ? input.zone : localZone());
				row.set('repeat', 'none');
				apply(row, { ...input, start, end });
				map.set(id, row);
			});
			return get(id);
		},
		update(id, patch) {
			const row = map.get(id);
			const current = get(id);
			if (!row || !current || current.organizerId !== user.id || current.cancelled) return null;
			const start = patch.start ?? current.start;
			const end = patch.end ?? current.end;
			if (end <= start || end - start > MAX_MEETING_HOURS * 3_600_000) return null;
			if (patch.title !== undefined && !sanitizeChannelName(patch.title)) return null;
			doc.transact(() => apply(row, patch));
			return get(id);
		},
		cancel(id) {
			const row = map.get(id);
			const current = get(id);
			if (!row || !current || current.organizerId !== user.id) return false;
			row.set('cancelled', true);
			return true;
		},
	};
}

/** The channel-style key a meeting's call and live flag use. */
export const meetingCallKey = (meetingId: string): string => `mtg-${meetingId}`.slice(0, 64);
