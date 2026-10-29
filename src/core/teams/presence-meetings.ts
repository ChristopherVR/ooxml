// Meeting-driven presence, as pure rules: the time windows in which a person is in a scheduled
// meeting, whether `now` is inside one and when that next changes (what the presence controller
// wakes up for). Only meetings the person is part of count: they organise it or are on the
// attendee list. A channel meeting also shows in everyone's calendar, but that does not mean
// everyone attends it, so it does not make the whole workspace Busy. Cancelled meetings and
// cancelled occurrences have no windows (`occurrencesBetween`). No clock: callers pass `now`.
// DOM-free. New code.
import { MINUTE_MS } from './time-zone';
import { occurrencesBetween } from './calendar';
import { type Meeting, meetingVisibleTo } from './meetings';

/** A span of time in a meeting, UTC milliseconds, `start` included and `end` not. */
export interface MeetingWindow {
	start: number;
	end: number;
}

/** How far ahead windows are listed. Longer than any meeting (24 hours), so nothing is cut short. */
export const MEETING_WINDOW_HORIZON_MS = 48 * 60 * MINUTE_MS;

/** Whether the person is expected in the meeting: organiser or on the attendee list. */
export const attendsMeeting = (meeting: Meeting, userId: string): boolean =>
	meetingVisibleTo(meeting, userId) &&
	(meeting.organizerId === userId || meeting.attendees.some((a) => a.id === userId));

/**
 * The person's meeting windows that end after `now` and start within the horizon, in order.
 * Overlapping and back-to-back meetings are merged, so presence does not flicker between them.
 */
export function meetingWindows(
	meetings: readonly Meeting[],
	userId: string,
	now: number,
	horizon: number = MEETING_WINDOW_HORIZON_MS,
): MeetingWindow[] {
	const spans: MeetingWindow[] = [];
	for (const meeting of meetings) {
		if (!attendsMeeting(meeting, userId)) continue;
		for (const o of occurrencesBetween(meeting, now, now + horizon))
			if (o.end > o.start) spans.push({ start: o.start, end: o.end });
	}
	spans.sort((a, b) => a.start - b.start || a.end - b.end);
	const merged: MeetingWindow[] = [];
	for (const span of spans) {
		const last = merged[merged.length - 1];
		if (last && span.start <= last.end) last.end = Math.max(last.end, span.end);
		else merged.push({ ...span });
	}
	return merged;
}

/** True while `now` is inside any window. */
export const inMeetingAt = (windows: readonly MeetingWindow[], now: number): boolean =>
	windows.some((w) => w.start <= now && now < w.end);

/** The next start or end after `now`, or null when no window has one left. */
export function nextMeetingBoundary(windows: readonly MeetingWindow[], now: number): number | null {
	let next: number | null = null;
	for (const w of windows)
		for (const t of [w.start, w.end]) if (t > now && (next === null || t < next)) next = t;
	return next;
}

export const sameWindows = (a: readonly MeetingWindow[], b: readonly MeetingWindow[]): boolean =>
	a.length === b.length && a.every((w, i) => w.start === b[i]!.start && w.end === b[i]!.end);
