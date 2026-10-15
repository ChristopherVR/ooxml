// The occurrence series of a repeating meeting, before exceptions: daily and weekly repeats keep
// the organiser's wall-clock time across daylight saving, and the series ends after `count`
// occurrences or at the instant `until` (an occurrence starting at or before it still happens).
// A moved or cancelled single occurrence is an exception, applied in `calendar.ts`. DOM-free.
// New code.
import type { Meeting } from './meetings';
import { DAY_MS, zonedParts, zonedToUtc } from './time-zone';

export const MAX_RECURRENCE_COUNT = 999;

type Series = Pick<Meeting, 'start' | 'end' | 'zone' | 'repeat' | 'until' | 'count'>;

export interface SeriesSlot {
	/** 0 for the first occurrence. */
	index: number;
	start: number;
}

const stepDays = (meeting: Pick<Meeting, 'repeat'>): number => (meeting.repeat === 'daily' ? 1 : 7);

/** The start of the `index`-th occurrence, ignoring the end of the series. */
export function seriesStart(meeting: Series, index: number): number {
	if (index === 0 || meeting.repeat === 'none') return meeting.start;
	const anchor = zonedParts(meeting.start, meeting.zone);
	return zonedToUtc({ ...anchor, day: anchor.day + index * stepDays(meeting) }, meeting.zone);
}

/** Whether the series still has an occurrence at this index and start. */
export function withinSeries(meeting: Series, index: number, start: number): boolean {
	if (index < 0) return false;
	if (meeting.repeat === 'none') return index === 0;
	if (meeting.count !== undefined && index >= meeting.count) return false;
	return meeting.until === undefined || start <= meeting.until;
}

/** Series occurrences that overlap `[from, to)`, in start order, at most `limit`. */
export function seriesBetween(
	meeting: Series,
	from: number,
	to: number,
	limit = 400,
): SeriesSlot[] {
	const span = meeting.end - meeting.start;
	if (to <= from) return [];
	if (meeting.repeat === 'none')
		return meeting.start < to && meeting.end > from ? [{ index: 0, start: meeting.start }] : [];
	const step = stepDays(meeting) * DAY_MS;
	const first = Math.max(0, Math.floor((from - meeting.end) / step) - 1);
	const last = Math.ceil((to - meeting.start) / step) + 1;
	const out: SeriesSlot[] = [];
	for (let k = first; k <= last && out.length < limit; k++) {
		const start = seriesStart(meeting, k);
		if (!withinSeries(meeting, k, start)) {
			if (meeting.count !== undefined && k >= meeting.count) break;
			if (meeting.until !== undefined && start > meeting.until) break;
			continue;
		}
		if (start < to && start + span > from) out.push({ index: k, start });
	}
	return out;
}

/** The index of the occurrence that starts exactly at `start`, or null when there is none. */
export function seriesIndexOf(meeting: Series, start: number): number | null {
	if (meeting.repeat === 'none') return start === meeting.start ? 0 : null;
	const guess = Math.round((start - meeting.start) / (stepDays(meeting) * DAY_MS));
	for (const k of [guess, guess - 1, guess + 1])
		if (k >= 0 && seriesStart(meeting, k) === start && withinSeries(meeting, k, start)) return k;
	return null;
}
