import * as Y from 'yjs';
import { createMeetingStore } from './meetings';
import {
	attendsMeeting,
	inMeetingAt,
	meetingWindows,
	nextMeetingBoundary,
	sameWindows,
} from './presence-meetings';

const HOUR = 3_600_000;
const T0 = Date.UTC(2026, 9, 20, 13);
const ada = { id: 'ada', name: 'Ada' };
const bob = { id: 'bob', name: 'Bob' };

function store(organizer = ada) {
	return createMeetingStore(new Y.Doc(), organizer, () => T0);
}

describe('meeting windows', () => {
	it('counts meetings the person organises or is invited to, not others', () => {
		const s = store();
		const mine = s.create({ title: 'Mine', start: T0, end: T0 + HOUR })!;
		const invited = store(bob).create({
			title: 'Invited',
			start: T0,
			end: T0 + HOUR,
			attendees: [ada],
		})!;
		const other = store(bob).create({ title: 'Other', start: T0, end: T0 + HOUR })!;
		expect(attendsMeeting(mine, 'ada')).toBe(true);
		expect(attendsMeeting(invited, 'ada')).toBe(true);
		expect(attendsMeeting(other, 'ada')).toBe(false);
		expect(meetingWindows([other], 'ada', T0)).toEqual([]);
	});

	it('does not count a channel meeting for people who are not on it', () => {
		const m = store(bob).create({
			title: 'Standup',
			start: T0,
			end: T0 + HOUR,
			channelId: 'general',
		})!;
		expect(attendsMeeting(m, 'ada')).toBe(false);
		expect(attendsMeeting(m, 'bob')).toBe(true);
	});

	it('skips cancelled meetings and cancelled occurrences', () => {
		const s = store();
		const m = s.create({ title: 'Gone', start: T0, end: T0 + HOUR })!;
		const series = s.create({ title: 'Daily', start: T0, end: T0 + HOUR, repeat: 'daily' })!;
		expect(meetingWindows(s.list(), 'ada', T0 - HOUR)).toHaveLength(2);
		expect(s.cancel(m.id)).toBe(true);
		expect(s.cancelOccurrence(series.id, T0)).toBe(true);
		expect(meetingWindows(s.list(), 'ada', T0 - HOUR, 12 * HOUR)).toEqual([]);
		expect(meetingWindows(s.list(), 'ada', T0 - HOUR, 30 * HOUR)).toEqual([
			{ start: T0 + 24 * HOUR, end: T0 + 25 * HOUR },
		]);
	});

	it('includes a meeting already in progress and merges overlapping and back-to-back ones', () => {
		const s = store();
		s.create({ title: 'A', start: T0 - HOUR, end: T0 + HOUR })!;
		s.create({ title: 'B', start: T0 + HOUR / 2, end: T0 + 2 * HOUR })!;
		s.create({ title: 'C', start: T0 + 2 * HOUR, end: T0 + 3 * HOUR })!;
		s.create({ title: 'D', start: T0 + 5 * HOUR, end: T0 + 6 * HOUR })!;
		const windows = meetingWindows(s.list(), 'ada', T0);
		expect(windows).toEqual([
			{ start: T0 - HOUR, end: T0 + 3 * HOUR },
			{ start: T0 + 5 * HOUR, end: T0 + 6 * HOUR },
		]);
		expect(sameWindows(windows, meetingWindows(s.list(), 'ada', T0))).toBe(true);
		expect(sameWindows(windows, windows.slice(1))).toBe(false);
	});

	it('answers whether now is inside a window (end excluded) and when that next changes', () => {
		const w = [
			{ start: 100, end: 200 },
			{ start: 300, end: 400 },
		];
		expect(inMeetingAt(w, 99)).toBe(false);
		expect(inMeetingAt(w, 100)).toBe(true);
		expect(inMeetingAt(w, 199)).toBe(true);
		expect(inMeetingAt(w, 200)).toBe(false);
		expect(nextMeetingBoundary(w, 0)).toBe(100);
		expect(nextMeetingBoundary(w, 100)).toBe(200);
		expect(nextMeetingBoundary(w, 250)).toBe(300);
		expect(nextMeetingBoundary(w, 400)).toBeNull();
		expect(nextMeetingBoundary([], 0)).toBeNull();
	});
});
