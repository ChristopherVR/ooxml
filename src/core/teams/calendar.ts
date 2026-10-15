// Pure calendar rules for the week view: which occurrences of a meeting fall in a range (daily and
// weekly repeats keep the organiser's wall-clock time), the days a work week / week / day view
// shows in the viewer's zone, how overlapping meetings share a day column and which reminders are
// due. No clock and no storage: callers pass `now`. DOM-free. New code.
import type { Meeting } from './meetings';
import {
	DAY_MS,
	MINUTE_MS,
	addDays,
	minuteOfDay,
	startOfDay,
	zonedParts,
	zonedToUtc,
} from './time-zone';

export interface Occurrence {
	/** `<meeting id>@<start>`: stable per occurrence, what reminders remember. */
	key: string;
	meetingId: string;
	start: number;
	end: number;
	/** 0 for the first occurrence. */
	index: number;
}

const MAX_OCCURRENCES = 400;

/** Occurrences that overlap `[from, to)`, in start order. Cancelled meetings have none. */
export function occurrencesBetween(meeting: Meeting, from: number, to: number): Occurrence[] {
	if (meeting.cancelled || to <= from) return [];
	const span = meeting.end - meeting.start;
	const make = (index: number, start: number): Occurrence => ({
		key: `${meeting.id}@${start}`,
		meetingId: meeting.id,
		start,
		end: start + span,
		index,
	});
	if (meeting.repeat === 'none')
		return meeting.start < to && meeting.end > from ? [make(0, meeting.start)] : [];
	const step = meeting.repeat === 'daily' ? 1 : 7;
	const anchor = zonedParts(meeting.start, meeting.zone);
	const first = Math.max(0, Math.floor((from - meeting.end) / (step * DAY_MS)) - 1);
	const last = Math.ceil((to - meeting.start) / (step * DAY_MS)) + 1;
	const out: Occurrence[] = [];
	for (let k = first; k <= last && out.length < MAX_OCCURRENCES; k++) {
		const start =
			k === 0 ? meeting.start : zonedToUtc({ ...anchor, day: anchor.day + k * step }, meeting.zone);
		if (start < to && start + span > from) out.push(make(k, start));
	}
	return out;
}

export type CalendarMode = 'workweek' | 'week' | 'day';

export interface CalendarDay {
	/** UTC instant of local midnight. */
	start: number;
	/** UTC instant of the next local midnight. */
	end: number;
	/** 0 = Sunday. */
	weekday: number;
	day: number;
	month: number;
	year: number;
}

const describe = (start: number, zone: string): CalendarDay => {
	const p = zonedParts(start, zone);
	return {
		start,
		end: addDays(start, 1, zone),
		weekday: p.weekday,
		day: p.day,
		month: p.month,
		year: p.year,
	};
};

/** The days a view shows around `anchor`: Monday to Friday, Monday to Sunday, or just that day. */
export function viewDays(anchor: number, mode: CalendarMode, zone: string): CalendarDay[] {
	const today = startOfDay(anchor, zone);
	if (mode === 'day') return [describe(today, zone)];
	const sinceMonday = (zonedParts(anchor, zone).weekday + 6) % 7;
	const monday = addDays(today, -sinceMonday, zone);
	const count = mode === 'workweek' ? 5 : 7;
	return Array.from({ length: count }, (_, i) => describe(addDays(monday, i, zone), zone));
}

/** Previous (-1) or next (1) period: a day for the day view, a week otherwise. */
export function shiftAnchor(
	anchor: number,
	mode: CalendarMode,
	direction: -1 | 1,
	zone: string,
): number {
	return addDays(anchor, direction * (mode === 'day' ? 1 : 7), zone);
}

export interface PlacedOccurrence {
	occurrence: Occurrence;
	/** Minutes after local midnight, clipped to the day. */
	startMinute: number;
	endMinute: number;
	/** The side-by-side column and how many columns the overlapping group has. */
	column: number;
	columns: number;
}

/** Position occurrences in one day's column; overlapping ones share the width. */
export function layoutDay(
	occurrences: readonly Occurrence[],
	day: Pick<CalendarDay, 'start' | 'end'>,
	zone: string,
): PlacedOccurrence[] {
	const items = occurrences
		.filter((o) => o.start < day.end && o.end > day.start)
		.map((occurrence) => ({
			occurrence,
			startMinute: occurrence.start <= day.start ? 0 : minuteOfDay(occurrence.start, zone),
			endMinute:
				occurrence.end >= day.end ? 24 * 60 : Math.max(minuteOfDay(occurrence.end, zone), 1),
			column: 0,
			columns: 1,
		}))
		.sort((a, b) => a.startMinute - b.startMinute || b.endMinute - a.endMinute);
	let group: typeof items = [];
	let groupEnd = -1;
	const close = (): void => {
		const columns = group.reduce((m, i) => Math.max(m, i.column + 1), 1);
		for (const i of group) i.columns = columns;
		group = [];
	};
	for (const item of items) {
		if (group.length && item.startMinute >= groupEnd) close();
		const taken = new Set(group.filter((g) => g.endMinute > item.startMinute).map((g) => g.column));
		while (taken.has(item.column)) item.column++;
		group.push(item);
		groupEnd = Math.max(groupEnd, item.endMinute);
	}
	close();
	return items;
}

export const REMINDER_LEAD_MS = 5 * MINUTE_MS;

/** Occurrences that start within `lead` of `now` (and have not started). */
export function remindersDue(
	meetings: readonly Meeting[],
	now: number,
	lead: number = REMINDER_LEAD_MS,
): { meeting: Meeting; occurrence: Occurrence }[] {
	const out: { meeting: Meeting; occurrence: Occurrence }[] = [];
	for (const meeting of meetings)
		for (const occurrence of occurrencesBetween(meeting, now, now + lead + 1))
			if (occurrence.start > now && occurrence.start - now <= lead)
				out.push({ meeting, occurrence });
	return out.sort((a, b) => a.occurrence.start - b.occurrence.start);
}
