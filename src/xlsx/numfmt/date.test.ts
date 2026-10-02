import { describe, expect, it } from 'vitest';
import { dateParts } from './date.js';
import { dateToSerial, serialToDate } from './index.js';

const utc = (y: number, m: number, d: number, h = 0, min = 0): Date =>
	new Date(Date.UTC(y, m - 1, d, h, min));

describe('serial dates (1900 system)', () => {
	it.each([
		[1, utc(1900, 1, 1)],
		[59, utc(1900, 2, 28)],
		[61, utc(1900, 3, 1)],
		[45000, utc(2023, 3, 15)],
		[45000.75, utc(2023, 3, 15, 18)],
		[2958465, utc(9999, 12, 31)],
	])('serial %d', (serial, date) => {
		expect(serialToDate(serial).getTime()).toBe(date.getTime());
		expect(dateToSerial(date)).toBeCloseTo(serial, 9);
	});

	it('maps the fictitious 1900-02-29 (serial 60) onto 1900-03-01', () => {
		expect(serialToDate(60).toISOString().slice(0, 10)).toBe('1900-03-01');
		expect(dateToSerial(utc(1900, 3, 1))).toBe(61);
	});

	it('serial 0 is 1899-12-31 (shown by Excel as 1900-01-00)', () => {
		expect(serialToDate(0).toISOString().slice(0, 10)).toBe('1899-12-31');
		expect(dateToSerial(utc(1899, 12, 31))).toBe(0);
		expect(dateParts(0, false, 0)).toMatchObject({ year: 1900, month: 1, day: 0, weekday: 6 });
		expect(dateParts(60, false, 0)).toMatchObject({ year: 1900, month: 2, day: 29, weekday: 3 });
	});
});

describe('serial dates (1904 system)', () => {
	it.each([
		[0, utc(1904, 1, 1)],
		[1462, utc(1908, 1, 2)],
		[45000.25, utc(2027, 3, 16, 6)],
	])('serial %d', (serial, date) => {
		expect(serialToDate(serial, true).getTime()).toBe(date.getTime());
		expect(dateToSerial(date, true)).toBeCloseTo(serial, 9);
	});
});

describe('dateParts rounding', () => {
	it('rounds to whole seconds and carries into the next day', () => {
		expect(dateParts(45000.999999, false, 0)).toMatchObject({ day: 16, hours: 0, seconds: 0 });
	});

	it('keeps sub-second digits when asked', () => {
		expect(dateParts(0.123456, false, 2)).toMatchObject({
			hours: 2,
			minutes: 57,
			seconds: 46,
			subsec: 60,
		});
	});
});
