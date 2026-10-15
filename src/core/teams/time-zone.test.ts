import {
	DAY_MS,
	addDays,
	dayKey,
	isValidZone,
	minuteOfDay,
	offsetMinutes,
	startOfDay,
	zonedParts,
	zonedToUtc,
} from './time-zone';

const NY = 'America/New_York';

describe('time zones', () => {
	it('validates zone names', () => {
		expect(isValidZone('Europe/Berlin')).toBe(true);
		expect(isValidZone('Mars/Olympus')).toBe(false);
		expect(isValidZone('')).toBe(false);
		expect(isValidZone(5)).toBe(false);
	});

	it('reads wall-clock parts and offsets', () => {
		const ms = Date.UTC(2026, 6, 1, 16, 30); // July: EDT, UTC-4
		const p = zonedParts(ms, NY);
		expect([p.year, p.month, p.day, p.hour, p.minute]).toEqual([2026, 7, 1, 12, 30]);
		expect(p.weekday).toBe(3);
		expect(offsetMinutes(ms, NY)).toBe(-240);
		expect(offsetMinutes(Date.UTC(2026, 0, 15), NY)).toBe(-300);
		expect(offsetMinutes(ms, 'UTC')).toBe(0);
		expect(offsetMinutes(ms, 'Asia/Kolkata')).toBe(330);
	});

	it('converts a wall time back to the instant', () => {
		expect(zonedToUtc({ year: 2026, month: 7, day: 1, hour: 12, minute: 30 }, NY)).toBe(
			Date.UTC(2026, 6, 1, 16, 30),
		);
		expect(zonedToUtc({ year: 2026, month: 1, day: 15, hour: 9 }, 'Europe/Berlin')).toBe(
			Date.UTC(2026, 0, 15, 8),
		);
	});

	it('moves a time in the spring-forward gap ahead and takes the first repeated hour', () => {
		// 2026-03-08 02:30 does not exist in New York; clocks jump from 02:00 to 03:00.
		const gap = zonedToUtc({ year: 2026, month: 3, day: 8, hour: 2, minute: 30 }, NY);
		expect(zonedParts(gap, NY).hour).toBe(3);
		// 2026-11-01 01:30 happens twice; the first is EDT.
		const twice = zonedToUtc({ year: 2026, month: 11, day: 1, hour: 1, minute: 30 }, NY);
		expect(twice).toBe(Date.UTC(2026, 10, 1, 5, 30));
	});

	it('keeps the local time of day across daylight saving', () => {
		const nineBefore = zonedToUtc({ year: 2026, month: 3, day: 7, hour: 9 }, NY);
		const nineAfter = addDays(nineBefore, 1, NY);
		expect(minuteOfDay(nineAfter, NY)).toBe(9 * 60);
		expect(nineAfter - nineBefore).toBe(DAY_MS - 3_600_000);
	});

	it('finds local midnight and keys days', () => {
		const ms = Date.UTC(2026, 6, 1, 3, 0); // still 30 June evening in New York
		expect(dayKey(ms, NY)).toBe('2026-06-30');
		expect(dayKey(ms, 'UTC')).toBe('2026-07-01');
		expect(startOfDay(ms, NY)).toBe(Date.UTC(2026, 5, 30, 4));
	});
});
