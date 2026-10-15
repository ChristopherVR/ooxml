import {
	layoutDay,
	occurrencesBetween,
	remindersDue,
	shiftAnchor,
	viewDays,
	type Occurrence,
} from './calendar';
import { icsEscape, icsFold, icsFileName, icsUtc, meetingToIcs } from './ics';
import type { Meeting } from './meetings';
import { DAY_MS, HOUR_MS, minuteOfDay, zonedToUtc } from './time-zone';

const NY = 'America/New_York';
const meeting = (over: Partial<Meeting> = {}): Meeting => ({
	id: 'm1',
	title: 'Stand-up',
	description: '',
	organizerId: 'ada',
	organizerName: 'Ada',
	zone: NY,
	start: zonedToUtc({ year: 2026, month: 3, day: 5, hour: 9 }, NY),
	end: zonedToUtc({ year: 2026, month: 3, day: 5, hour: 9, minute: 30 }, NY),
	repeat: 'none',
	attendees: [],
	createdAt: 0,
	cancelled: false,
	...over,
});

describe('occurrences', () => {
	it('returns a single meeting only inside the range', () => {
		const m = meeting();
		expect(occurrencesBetween(m, m.start - HOUR_MS, m.end + HOUR_MS)).toHaveLength(1);
		expect(occurrencesBetween(m, m.end, m.end + HOUR_MS)).toHaveLength(0);
		expect(occurrencesBetween(m, m.start - HOUR_MS, m.start)).toHaveLength(0);
		expect(occurrencesBetween(meeting({ cancelled: true }), 0, 1e15)).toEqual([]);
	});

	it('repeats daily at the organizer wall time across the spring-forward change', () => {
		const m = meeting({ repeat: 'daily' });
		const week = occurrencesBetween(m, m.start, m.start + 6.5 * DAY_MS);
		expect(week).toHaveLength(7);
		for (const o of week) expect(minuteOfDay(o.start, NY)).toBe(9 * 60);
		// 8 March 2026 is the New York DST change: that day is one hour shorter.
		expect(week[3]!.start - week[2]!.start).toBe(DAY_MS - HOUR_MS);
		expect(week.every((o) => o.end - o.start === 30 * 60_000)).toBe(true);
	});

	it('repeats weekly on the same weekday and never before its first start', () => {
		const m = meeting({ repeat: 'weekly' });
		const month = occurrencesBetween(m, m.start - 30 * DAY_MS, m.start + 28 * DAY_MS);
		expect(month.map((o) => o.index)).toEqual([0, 1, 2, 3, 4]);
		expect(month[1]!.start - month[0]!.start).toBeGreaterThanOrEqual(7 * DAY_MS - HOUR_MS);
		expect(new Set(month.map((o) => o.key)).size).toBe(5);
	});

	it('finds a late occurrence without walking from the start', () => {
		const m = meeting({ repeat: 'daily' });
		const far = m.start + 1000 * DAY_MS;
		const found = occurrencesBetween(m, far, far + DAY_MS);
		expect(found.length).toBeGreaterThanOrEqual(1);
		expect(found.length).toBeLessThanOrEqual(2);
	});
});

describe('week view', () => {
	const thursday = zonedToUtc({ year: 2026, month: 10, day: 15, hour: 14 }, NY);
	it('shows Monday to Friday, Monday to Sunday or one day in the viewer zone', () => {
		const work = viewDays(thursday, 'workweek', NY);
		expect(work.map((d) => d.day)).toEqual([12, 13, 14, 15, 16]);
		expect(viewDays(thursday, 'week', NY).map((d) => d.weekday)).toEqual([1, 2, 3, 4, 5, 6, 0]);
		expect(viewDays(thursday, 'day', NY).map((d) => d.day)).toEqual([15]);
		expect(work[0]!.end).toBe(work[1]!.start);
	});

	it('steps by a day or a week', () => {
		expect(viewDays(shiftAnchor(thursday, 'week', 1, NY), 'week', NY)[0]!.day).toBe(19);
		expect(viewDays(shiftAnchor(thursday, 'day', -1, NY), 'day', NY)[0]!.day).toBe(14);
	});

	it('renders a meeting in the viewer zone, not the organizer zone', () => {
		const m = meeting({ zone: 'Asia/Tokyo' });
		const day = viewDays(m.start, 'day', NY)[0]!;
		const [placed] = layoutDay(occurrencesBetween(m, day.start, day.end), day, NY);
		expect(placed!.startMinute).toBe(minuteOfDay(m.start, NY));
		expect(placed!.endMinute - placed!.startMinute).toBe(30);
	});
});

describe('day layout', () => {
	const day = { start: 0, end: DAY_MS };
	const occ = (id: string, from: number, to: number): Occurrence => ({
		key: id,
		meetingId: id,
		start: from * 60_000,
		end: to * 60_000,
		index: 0,
	});
	it('puts overlapping meetings side by side and lets later ones reuse the width', () => {
		const placed = layoutDay(
			[occ('a', 540, 600), occ('b', 570, 630), occ('c', 600, 660), occ('d', 720, 780)],
			day,
			'UTC',
		);
		const by = Object.fromEntries(placed.map((p) => [p.occurrence.key, p]));
		expect([by.a!.column, by.b!.column]).toEqual([0, 1]);
		expect(by.c!.column).toBe(0);
		expect([by.a!.columns, by.b!.columns, by.c!.columns]).toEqual([2, 2, 2]);
		expect([by.d!.column, by.d!.columns]).toEqual([0, 1]);
	});

	it('clips a meeting that crosses midnight to each day', () => {
		const late = occ('late', 23 * 60, 25 * 60);
		const [first] = layoutDay([late], day, 'UTC');
		expect([first!.startMinute, first!.endMinute]).toEqual([1380, 1440]);
		const [second] = layoutDay([late], { start: DAY_MS, end: 2 * DAY_MS }, 'UTC');
		expect([second!.startMinute, second!.endMinute]).toEqual([0, 60]);
	});
});

describe('reminders', () => {
	it('is due in the five minutes before an occurrence starts', () => {
		const m = meeting({ repeat: 'daily' });
		const next = m.start + DAY_MS;
		expect(remindersDue([m], next - 6 * 60_000)).toEqual([]);
		const due = remindersDue([m], next - 4 * 60_000);
		expect(due.map((d) => d.occurrence.start)).toEqual([
			zonedToUtc({ year: 2026, month: 3, day: 6, hour: 9 }, NY),
		]);
		expect(remindersDue([m], next + 1)).toEqual([]);
		expect(remindersDue([meeting({ cancelled: true })], m.start - 60_000)).toEqual([]);
	});
});

describe('ics', () => {
	it('escapes, folds and writes UTC times', () => {
		expect(icsUtc(Date.UTC(2026, 9, 15, 14, 5, 9))).toBe('20261015T140509Z');
		expect(icsEscape('a,b;c\\d\ne')).toBe('a\\,b\\;c\\\\d\\ne');
		const folded = icsFold(`SUMMARY:${'x'.repeat(200)}`);
		expect(folded.length).toBeGreaterThan(2);
		expect(folded.every((l) => new TextEncoder().encode(l).length <= 75)).toBe(true);
		expect(folded.map((l, i) => (i ? l.slice(1) : l)).join('')).toBe(`SUMMARY:${'x'.repeat(200)}`);
	});

	it('exports a meeting with its repeat and a five minute alarm', () => {
		const text = meetingToIcs(
			meeting({
				repeat: 'weekly',
				description: 'Agenda, notes',
				attendees: [{ id: 'bob', name: 'Bob' }],
			}),
			{ now: Date.UTC(2026, 0, 1) },
		);
		expect(text.endsWith('\r\n')).toBe(true);
		expect(text.split('\r\n')).toEqual(
			expect.arrayContaining([
				'BEGIN:VCALENDAR',
				'UID:m1@openteams',
				'DTSTART:20260305T140000Z',
				'DTEND:20260305T143000Z',
				'SUMMARY:Stand-up',
				'RRULE:FREQ=WEEKLY',
				'TRIGGER:-PT5M',
				'END:VCALENDAR',
			]),
		);
		expect(text).toContain('Agenda\\, notes');
		expect(text).toContain('Attendees: Bob');
		expect(icsFileName({ title: 'Q4 / plan: draft' })).toBe('Q4-plan-draft.ics');
	});
});
