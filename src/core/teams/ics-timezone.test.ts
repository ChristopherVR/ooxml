import { icsDateTime, icsLocal, meetingToIcs } from './ics';
import { icsOffset, vtimezoneLines, zoneTransitions } from './ics-timezone';
import type { Meeting } from './meetings';
import { HOUR_MS, offsetMinutes, zonedToUtc } from './time-zone';

describe('zone transitions', () => {
	it('finds the two United States changes of 2026 to the minute', () => {
		expect(zoneTransitions('America/New_York', 2026)).toEqual([
			{ at: Date.UTC(2026, 2, 8, 7), from: -300, to: -240 },
			{ at: Date.UTC(2026, 10, 1, 6), from: -240, to: -300 },
		]);
	});

	it('finds none in a zone without daylight saving', () => {
		expect(zoneTransitions('Asia/Kolkata', 2026)).toEqual([]);
		expect(zoneTransitions('UTC', 2026)).toEqual([]);
	});

	it('formats offsets with half hours and negative values', () => {
		expect(icsOffset(330)).toBe('+0530');
		expect(icsOffset(-210)).toBe('-0330');
		expect(icsOffset(0)).toBe('+0000');
	});

	it('matches the platform offsets on both sides of a change', () => {
		for (const c of zoneTransitions('Europe/Berlin', 2027)) {
			expect(offsetMinutes(c.at - HOUR_MS, 'Europe/Berlin')).toBe(c.from);
			expect(offsetMinutes(c.at, 'Europe/Berlin')).toBe(c.to);
		}
	});
});

describe('VTIMEZONE', () => {
	it('writes yearly rules for the United States', () => {
		const lines = vtimezoneLines('America/New_York', 2026);
		expect(lines[0]).toBe('BEGIN:VTIMEZONE');
		expect(lines).toContain('TZID:America/New_York');
		expect(lines).toContain('RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU');
		expect(lines).toContain('RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU');
		expect(lines).toContain('DTSTART:20250309T020000');
		expect(lines).toContain('DTSTART:20251102T020000');
		expect(lines).toContain('TZOFFSETFROM:-0500');
		expect(lines).toContain('TZOFFSETTO:-0400');
		expect(lines.at(-1)).toBe('END:VTIMEZONE');
	});

	it('uses the last Sunday for the European rule', () => {
		const lines = vtimezoneLines('Europe/Berlin', 2026);
		expect(lines).toContain('RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU');
		expect(lines).toContain('RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU');
		expect(lines).toContain('DTSTART:20250330T020000');
		expect(lines).toContain('DTSTART:20251026T030000');
		expect(lines).toContain('TZOFFSETTO:+0200');
	});

	it('orders the southern hemisphere by date', () => {
		const lines = vtimezoneLines('Australia/Sydney', 2026);
		expect(lines).toContain('RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=1SU');
		expect(lines).toContain('RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=1SU');
		expect(lines).toContain('TZOFFSETTO:+1100');
	});

	it('writes one fixed offset where the clocks never change', () => {
		const lines = vtimezoneLines('Asia/Kolkata', 2026);
		expect(lines).toContain('TZOFFSETFROM:+0530');
		expect(lines).toContain('TZOFFSETTO:+0530');
		expect(lines.some((l) => l.startsWith('RRULE'))).toBe(false);
		expect(lines.filter((l) => l.startsWith('BEGIN:'))).toEqual([
			'BEGIN:VTIMEZONE',
			'BEGIN:STANDARD',
		]);
	});

	it('lists the changes one by one where the rules stop repeating', () => {
		// Brazil gave up daylight saving after February 2019.
		const lines = vtimezoneLines('America/Sao_Paulo', 2019);
		expect(lines.some((l) => l.startsWith('RRULE'))).toBe(false);
		expect(lines).toContain('BEGIN:DAYLIGHT');
		expect(lines).toContain('TZOFFSETTO:-0300');
	});

	it('has balanced BEGIN and END lines', () => {
		for (const zone of ['America/New_York', 'Europe/London', 'Asia/Tokyo', 'America/Sao_Paulo']) {
			const open: string[] = [];
			for (const l of vtimezoneLines(zone, 2026)) {
				if (l.startsWith('BEGIN:')) open.push(l.slice(6));
				else if (l.startsWith('END:')) expect(open.pop()).toBe(l.slice(4));
			}
			expect(open).toEqual([]);
		}
	});
});

describe('zoned date-times', () => {
	it('writes wall time in the zone, or UTC for a UTC organizer', () => {
		const ms = zonedToUtc({ year: 2026, month: 10, day: 15, hour: 9 }, 'Europe/Berlin');
		expect(icsLocal(ms, 'Europe/Berlin')).toBe('20261015T090000');
		expect(icsDateTime('DTSTART', ms, 'Europe/Berlin')).toBe(
			'DTSTART;TZID=Europe/Berlin:20261015T090000',
		);
		expect(icsDateTime('DTSTART', ms, 'UTC')).toBe('DTSTART:20261015T070000Z');
	});
});

describe('exported recurring meeting', () => {
	const NY = 'America/New_York';
	const at = (month: number, day: number, hour = 9): number =>
		zonedToUtc({ year: 2026, month, day, hour }, NY);
	const base: Meeting = {
		id: 'm1',
		title: 'Weekly, sync; "team"',
		description: 'Line one\nLine two',
		organizerId: 'ada',
		organizerName: 'Ada',
		zone: NY,
		start: at(3, 2),
		end: at(3, 2) + HOUR_MS,
		repeat: 'weekly',
		attendees: [],
		createdAt: 0,
		cancelled: false,
	};
	const lines = (m: Meeting): string[] =>
		meetingToIcs(m, { now: Date.UTC(2026, 0, 1) })
			.replace(/\r\n /gu, '')
			.split('\r\n');

	it('puts the VTIMEZONE before the event and keeps 09:00 local', () => {
		const out = lines(base);
		expect(out.indexOf('BEGIN:VTIMEZONE')).toBeGreaterThan(out.indexOf('BEGIN:VCALENDAR'));
		expect(out.indexOf('END:VTIMEZONE')).toBeLessThan(out.indexOf('BEGIN:VEVENT'));
		expect(out).toContain('DTSTART;TZID=America/New_York:20260302T090000');
		expect(out).toContain('DTEND;TZID=America/New_York:20260302T100000');
		expect(out).toContain('RRULE:FREQ=WEEKLY');
		expect(out).toContain('SUMMARY:Weekly\\, sync\\; "team"');
		expect(out).toContain('DESCRIPTION:Line one\\nLine two\\n\\nOrganizer: Ada');
		expect(out.some((l) => l.startsWith('DTSTART') && l.endsWith('Z'))).toBe(false);
	});

	it('writes COUNT or UNTIL (in UTC) and EXDATE in the zone', () => {
		expect(lines({ ...base, count: 6 })).toContain('RRULE:FREQ=WEEKLY;COUNT=6');
		const until = lines({ ...base, until: Date.UTC(2026, 5, 1, 3, 59, 59) });
		expect(until).toContain('RRULE:FREQ=WEEKLY;UNTIL=20260601T035959Z');
		const ex = lines({
			...base,
			exceptions: [
				{ originalStart: at(3, 9), cancelled: true },
				{ originalStart: at(3, 16), cancelled: true },
			],
		});
		expect(ex).toContain('EXDATE;TZID=America/New_York:20260309T090000,20260316T090000');
	});

	it('writes a changed occurrence as an overriding event with RECURRENCE-ID', () => {
		const out = lines({
			...base,
			exceptions: [
				{
					originalStart: at(3, 9),
					cancelled: false,
					start: at(3, 9, 14),
					end: at(3, 9, 15),
					title: 'Moved',
				},
			],
		});
		expect(out.filter((l) => l === 'BEGIN:VEVENT')).toHaveLength(2);
		expect(out.filter((l) => l === 'UID:m1@openteams')).toHaveLength(2);
		const second = out.slice(out.lastIndexOf('BEGIN:VEVENT'));
		expect(second).toContain('RECURRENCE-ID;TZID=America/New_York:20260309T090000');
		expect(second).toContain('DTSTART;TZID=America/New_York:20260309T140000');
		expect(second).toContain('SUMMARY:Moved');
		expect(second.some((l) => l.startsWith('RRULE'))).toBe(false);
		expect(out.filter((l) => l.startsWith('EXDATE'))).toEqual([]);
	});

	it('uses plain UTC times and no VTIMEZONE for a UTC organizer', () => {
		const out = lines({
			...base,
			zone: 'UTC',
			exceptions: [{ originalStart: base.start, cancelled: true }],
		});
		expect(out).toContain('DTSTART:20260302T140000Z');
		expect(out).not.toContain('BEGIN:VTIMEZONE');
		expect(out).toContain('EXDATE:20260302T140000Z');
	});

	it('folds every line to 75 octets and ends each with CRLF', () => {
		const title = 'Ünïcödé '.repeat(40);
		const long = meetingToIcs({ ...base, title, description: 'é'.repeat(300) });
		const raw = long.split('\r\n');
		expect(raw.at(-1)).toBe('');
		const enc = new TextEncoder();
		for (const l of raw) expect(enc.encode(l).length).toBeLessThanOrEqual(75);
		expect(long).not.toMatch(/[^\r]\n/u);
		expect(long.replace(/\r\n /gu, '')).toContain(`SUMMARY:${title}`);
	});
});
