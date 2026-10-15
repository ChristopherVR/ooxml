// Scheduled meetings in the workspace Yjs document. A meeting stores its start and end as UTC
// instants plus the organiser's IANA zone, which only matters for repeats (a daily 09:00 stays 09:00
// on the organiser's clock across daylight saving); every viewer renders in their own zone. Rows
// are re-validated when read back. Any member can read the shared document, so the attendee list
// scopes who sees a meeting in their calendar, not who could read the row. DOM-free. New code.
import * as Y from 'yjs';
import { createIdGenerator } from '../collab/identity';
import { type ChatMember, isChatUserId, sanitizeMembers } from './chats';
import { sanitizeChannelName, sanitizeMessageText } from './model';
import {
	type MeetingException,
	type OccurrencePatch,
	exceptionAt,
	exceptionMap,
	exceptionValue,
	readExceptions,
} from './meeting-exceptions';
import { MAX_RECURRENCE_COUNT, seriesIndexOf } from './recurrence';
import {
	type MonthlyBy,
	WEEKDAY_ORDINALS,
	dayMatchesOrdinal,
	ordinalOfDay,
} from './recurrence-calendar';
import { isValidZone, localZone, zonedParts } from './time-zone';

export type { MeetingException, OccurrencePatch } from './meeting-exceptions';

export type MeetingRepeat = 'none' | 'daily' | 'weekly' | 'monthly' | 'yearly';
export const MEETING_REPEATS: readonly MeetingRepeat[] = [
	'none',
	'daily',
	'weekly',
	'monthly',
	'yearly',
];

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
	/**
	 * Monthly repeats only: on the same date (`date`, the default) or on the same weekday by
	 * ordinal (`weekday`, "the third Tuesday"). Yearly repeats are always on the date.
	 */
	monthlyBy?: MonthlyBy;
	/** With `monthlyBy: 'weekday'`: 1 to 4, or -1 for the last. Defaults to the start's own. */
	ordinal?: number;
	/** The series ends once an occurrence would start after this instant. Never with `count`. */
	until?: number;
	/** The series has this many occurrences in all. Never with `until`. */
	count?: number;
	/** Single occurrences that were cancelled or changed. */
	exceptions?: MeetingException[];
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
	monthlyBy?: MonthlyBy;
	ordinal?: number;
	/** Last instant a repeat may start at; null clears the end. */
	until?: number | null;
	/** Total occurrences of a repeat; null clears the end. */
	count?: number | null;
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
	/** Organiser only. Cancels one occurrence of a repeating meeting, keyed by its original start. */
	cancelOccurrence: (id: string, originalStart: number) => boolean;
	/** Organiser only. Changes one occurrence; the rest of the series is untouched. */
	editOccurrence: (id: string, originalStart: number, patch: OccurrencePatch) => Meeting | null;
}

const finiteTime = (v: unknown): number | null =>
	typeof v === 'number' && Number.isFinite(v) && v > 0 && v < 8.64e15 ? Math.floor(v) : null;

/**
 * The monthly pattern of a row. A weekday rule whose ordinal does not fit the start date (a
 * corrupt or hand-edited row) falls back to the ordinal the start date has.
 */
function readPattern(
	kind: MeetingRepeat,
	by: unknown,
	ordinal: unknown,
	start: number,
	zone: string,
): Pick<Meeting, 'monthlyBy' | 'ordinal'> {
	if (kind !== 'monthly' || by !== 'weekday') return {};
	const p = zonedParts(start, zone);
	const fits =
		typeof ordinal === 'number' &&
		WEEKDAY_ORDINALS.includes(ordinal) &&
		dayMatchesOrdinal(p.year, p.month, p.day, ordinal);
	return { monthlyBy: 'weekday', ordinal: fits ? (ordinal as number) : ordinalOfDay(p.day) };
}

/**
 * Validate the weekday ordinal of a monthly weekday rule against the start date, filling in the
 * start's own ordinal when none is given or the old one no longer fits. Null when an explicit
 * ordinal does not fit ("the third Tuesday" for a start that is not one).
 */
function withOrdinal<T extends Partial<MeetingInput>>(
	input: T,
	repeat: MeetingRepeat,
	by: MonthlyBy | undefined,
	current: number | undefined,
	start: number,
	zone: string,
): T | null {
	if (repeat !== 'monthly' || by !== 'weekday') return input;
	const p = zonedParts(start, zone);
	const fits = (o: number | undefined): o is number =>
		o !== undefined && WEEKDAY_ORDINALS.includes(o) && dayMatchesOrdinal(p.year, p.month, p.day, o);
	if (input.ordinal !== undefined) return fits(input.ordinal) ? input : null;
	return { ...input, ordinal: fits(current) ? current : ordinalOfDay(p.day) };
}

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
	const kind = MEETING_REPEATS.includes(repeat as MeetingRepeat)
		? (repeat as MeetingRepeat)
		: 'none';
	const count = raw.get('count');
	const until = finiteTime(raw.get('until'));
	const total =
		typeof count === 'number' && Number.isInteger(count) && count >= 1
			? Math.min(count, MAX_RECURRENCE_COUNT)
			: null;
	const exceptions = kind === 'none' ? [] : readExceptions(raw.get('exceptions'));
	const zoneId = isValidZone(zone) ? zone : 'UTC';
	const pattern = readPattern(kind, raw.get('monthlyBy'), raw.get('ordinal'), start, zoneId);
	return {
		id,
		title,
		description: sanitizeMessageText(raw.get('description')).slice(0, MAX_MEETING_DESCRIPTION),
		organizerId,
		organizerName: sanitizeChannelName(raw.get('organizerName')) || 'Unknown',
		zone: zoneId,
		start,
		end,
		repeat: kind,
		...pattern,
		...(kind === 'none' ? {} : total !== null ? { count: total } : until !== null ? { until } : {}),
		...(exceptions.length ? { exceptions } : {}),
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
		if (input.monthlyBy !== undefined) {
			if (input.monthlyBy === 'weekday' || input.monthlyBy === 'date')
				row.set('monthlyBy', input.monthlyBy);
		}
		if (input.ordinal !== undefined && WEEKDAY_ORDINALS.includes(input.ordinal))
			row.set('ordinal', input.ordinal);
		if (input.count !== undefined) {
			if (input.count === null) row.delete('count');
			else {
				row.set('count', Math.min(Math.max(1, Math.floor(input.count)), MAX_RECURRENCE_COUNT));
				row.delete('until');
			}
		}
		if (input.until !== undefined) {
			const until = input.until === null ? null : finiteTime(input.until);
			if (until === null) row.delete('until');
			else {
				row.set('until', until);
				row.delete('count');
			}
		}
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
			const zone = isValidZone(input.zone) ? input.zone : localZone();
			const repeat = input.repeat ?? 'none';
			const checked = withOrdinal(input, repeat, input.monthlyBy, undefined, start, zone);
			if (!checked) return null;
			const id = nextId('t').replace(/[^\w-]/gu, '');
			doc.transact(() => {
				const row = new Y.Map<unknown>();
				row.set('organizerId', user.id);
				row.set('organizerName', user.name);
				row.set('createdAt', now());
				row.set('cancelled', false);
				row.set('zone', zone);
				row.set('repeat', 'none');
				apply(row, { ...checked, start, end });
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
			const repeat = patch.repeat ?? current.repeat;
			const zone = patch.zone !== undefined && isValidZone(patch.zone) ? patch.zone : current.zone;
			const by = patch.monthlyBy ?? current.monthlyBy;
			const checked = withOrdinal(patch, repeat, by, current.ordinal, start, zone);
			if (!checked) return null;
			patch = checked;
			const sameShape =
				(by ?? 'date') === (current.monthlyBy ?? 'date') &&
				(patch.ordinal === undefined || patch.ordinal === current.ordinal);
			if (typeof patch.until === 'number' && repeat !== 'none' && patch.until < start) return null;
			doc.transact(() => {
				apply(row, patch);
				// Exceptions are keyed by the original start, so a new schedule voids them.
				if (repeat === 'none' || start !== current.start || repeat !== current.repeat || !sameShape)
					row.delete('exceptions');
				if (repeat === 'none') {
					row.delete('until');
					row.delete('count');
				}
			});
			return get(id);
		},
		cancel(id) {
			const row = map.get(id);
			const current = get(id);
			if (!row || !current || current.organizerId !== user.id) return false;
			row.set('cancelled', true);
			return true;
		},
		cancelOccurrence(id, originalStart) {
			const row = map.get(id);
			const current = get(id);
			if (!row || !current || current.organizerId !== user.id || current.cancelled) return false;
			if (seriesIndexOf(current, originalStart) === null || current.repeat === 'none') return false;
			doc.transact(() => exceptionMap(row).set(String(originalStart), { cancelled: true }));
			return true;
		},
		editOccurrence(id, originalStart, patch) {
			const row = map.get(id);
			const current = get(id);
			if (!row || !current || current.organizerId !== user.id || current.cancelled) return null;
			if (seriesIndexOf(current, originalStart) === null || current.repeat === 'none') return null;
			if (patch.title !== undefined && !sanitizeChannelName(patch.title)) return null;
			const previous = exceptionAt(current.exceptions, originalStart);
			const start = patch.start ?? previous?.start;
			const end = patch.end ?? previous?.end;
			if ((start === undefined) !== (end === undefined)) return null;
			if (start !== undefined && end !== undefined) {
				const finite = finiteTime(start) !== null && finiteTime(end) !== null;
				if (!finite || end <= start || end - start > MAX_MEETING_HOURS * 3_600_000) return null;
			}
			const value = exceptionValue(
				{ ...patch, ...(start !== undefined && end !== undefined ? { start, end } : {}) },
				previous,
			);
			doc.transact(() => exceptionMap(row).set(String(originalStart), value));
			return get(id);
		},
	};
}

/** The channel-style key a meeting's call and live flag use. */
export const meetingCallKey = (meetingId: string): string => `mtg-${meetingId}`.slice(0, 64);
