// Monthly and yearly repeats: the occurrences are calendar dates, not a fixed number of days, so they
// are generated month by month (or year by year) from the first occurrence, on the organiser's
// wall clock. Three shapes, as in Outlook: monthly on a date ("the 15th"), monthly on a weekday
// ordinal ("the third Tuesday", "the last Friday") and yearly on a date. A date that a month does
// not have (the 29th to 31st; 29 February in a common year) is skipped, not moved to the month's
// end, and a skipped month does not count towards `count`. DOM-free. New code.
import type { Meeting } from './meetings';
import { zonedParts, zonedToUtc } from './time-zone';

export type MonthlyBy = 'date' | 'weekday';

/** Ordinals a weekday rule may use: the first to fourth, or the last (-1). */
export const WEEKDAY_ORDINALS: readonly number[] = [1, 2, 3, 4, -1];

/** Months or years scanned for occurrences, which bounds a series far beyond any real calendar. */
const MAX_MONTH_STEPS = 24_000;
const MAX_YEAR_STEPS = 6_000;

type CalendarSeries = Pick<Meeting, 'start' | 'zone' | 'repeat' | 'monthlyBy' | 'ordinal'>;

export interface CalendarSlot {
	/** 0 for the first occurrence. Skipped months do not take an index. */
	index: number;
	/** The wall-clock date in the organiser's zone (the time of day is the series start's). */
	year: number;
	month: number;
	day: number;
}

export const isCalendarRepeat = (repeat: Meeting['repeat']): boolean =>
	repeat === 'monthly' || repeat === 'yearly';

export const daysInMonth = (year: number, month: number): number =>
	new Date(Date.UTC(year, month, 0)).getUTCDate();

/** The ordinal of a day of the month for its weekday: 1 to 4, and -1 (last) for the fifth. */
export const ordinalOfDay = (day: number): number => (day > 28 ? -1 : Math.ceil(day / 7));

/** Whether `day` is the `ordinal`-th weekday of its month, by that ordinal's own meaning. */
export function dayMatchesOrdinal(year: number, month: number, day: number, ordinal: number) {
	return ordinal === -1
		? day + 7 > daysInMonth(year, month)
		: Math.ceil(day / 7) === ordinal && ordinal >= 1 && ordinal <= 4;
}

/** The day of the month on which weekday `weekday` (0 = Sunday) falls for this ordinal. */
export function nthWeekdayOfMonth(
	year: number,
	month: number,
	weekday: number,
	ordinal: number,
): number {
	const first = new Date(Date.UTC(year, month - 1, 1)).getUTCDay();
	if (ordinal === -1) {
		const length = daysInMonth(year, month);
		const lastWeekday = (first + length - 1) % 7;
		return length - ((lastWeekday - weekday + 7) % 7);
	}
	return 1 + ((weekday - first + 7) % 7) + (ordinal - 1) * 7;
}

/** The anchor (first occurrence) as a wall-clock date, and the rule derived from it. */
function anchorOf(meeting: CalendarSeries) {
	const p = zonedParts(meeting.start, meeting.zone);
	const weekday = new Date(Date.UTC(p.year, p.month - 1, p.day)).getUTCDay();
	const ordinal =
		meeting.ordinal !== undefined && WEEKDAY_ORDINALS.includes(meeting.ordinal)
			? meeting.ordinal
			: ordinalOfDay(p.day);
	return { ...p, weekday, ordinal, byWeekday: meeting.monthlyBy === 'weekday' };
}

/** The date of step `j` (0 = the first), or null when that month or year has no such date. */
function candidate(
	a: ReturnType<typeof anchorOf>,
	yearly: boolean,
	j: number,
): { year: number; month: number; day: number } | null {
	if (j === 0) return { year: a.year, month: a.month, day: a.day };
	if (yearly) {
		const year = a.year + j;
		return a.day <= daysInMonth(year, a.month) ? { year, month: a.month, day: a.day } : null;
	}
	const total = a.year * 12 + (a.month - 1) + j;
	const year = Math.floor(total / 12);
	const month = (total % 12) + 1;
	if (a.byWeekday)
		return { year, month, day: nthWeekdayOfMonth(year, month, a.weekday, a.ordinal) };
	return a.day <= daysInMonth(year, month) ? { year, month, day: a.day } : null;
}

/** Every date of a monthly or yearly series, in order, until the scan limit. */
export function* calendarSlots(meeting: CalendarSeries): Generator<CalendarSlot> {
	const a = anchorOf(meeting);
	const yearly = meeting.repeat === 'yearly';
	let index = 0;
	for (let j = 0; j < (yearly ? MAX_YEAR_STEPS : MAX_MONTH_STEPS); j++) {
		const d = candidate(a, yearly, j);
		if (d) yield { index: index++, ...d };
	}
}

/** The instant of a slot: the series' time of day on that date in the organiser's zone. */
export function slotStart(meeting: CalendarSeries, slot: CalendarSlot): number {
	const p = zonedParts(meeting.start, meeting.zone);
	return zonedToUtc({ ...slot, hour: p.hour, minute: p.minute, second: p.second }, meeting.zone);
}

/** The slot at `index`, or null when the series has fewer dates within the scan limit. */
export function slotAt(meeting: CalendarSeries, index: number): CalendarSlot | null {
	for (const slot of calendarSlots(meeting)) if (slot.index === index) return slot;
	return null;
}

/** The slot that starts exactly at `start`, or null. */
export function slotStarting(meeting: CalendarSeries, start: number): CalendarSlot | null {
	const p = zonedParts(start, meeting.zone);
	for (const slot of calendarSlots(meeting)) {
		if (slot.year > p.year || (slot.year === p.year && slot.month > p.month)) return null;
		if (slot.year === p.year && slot.month === p.month && slot.day === p.day)
			return slotStart(meeting, slot) === start ? slot : null;
	}
	return null;
}

/** The calendar facts of a start instant in a zone that a repeat picker needs. */
export interface RepeatChoices {
	month: number;
	day: number;
	/** 0 = Sunday. */
	weekday: number;
	/** The ordinals that fit this date, e.g. `[4, -1]` for a Friday that is fourth and last. */
	ordinals: number[];
}

export function repeatChoices(start: number, zone: string): RepeatChoices {
	const p = zonedParts(start, zone);
	const weekday = new Date(Date.UTC(p.year, p.month - 1, p.day)).getUTCDay();
	const ordinals = WEEKDAY_ORDINALS.filter((o) => dayMatchesOrdinal(p.year, p.month, p.day, o));
	return { month: p.month, day: p.day, weekday, ordinals };
}
