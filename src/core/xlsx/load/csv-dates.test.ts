import { describe, expect, it } from 'vitest';
import { dateToSerial } from '../numfmt/date';
import { parseOrderedDate } from './csv-dates';

const at = (year: number, month: number, day: number, h = 0, m = 0, s = 0): number =>
	dateToSerial(new Date(Date.UTC(year, month - 1, day, h, m, s)));

describe('parseOrderedDate', () => {
	it('reads the parts in each Text Import Wizard order', () => {
		const march4 = at(2024, 3, 4);
		expect(parseOrderedDate('04/03/2024', 'MDY', false)?.value).toBe(at(2024, 4, 3));
		expect(parseOrderedDate('04/03/2024', 'DMY', false)?.value).toBe(march4);
		expect(parseOrderedDate('2024/03/04', 'YMD', false)?.value).toBe(march4);
		expect(parseOrderedDate('03/2024/04', 'MYD', false)?.value).toBe(march4);
		expect(parseOrderedDate('04/2024/03', 'DYM', false)?.value).toBe(march4);
		expect(parseOrderedDate('2024/04/03', 'YDM', false)?.value).toBe(march4);
		expect(parseOrderedDate('04/03/2024', 'DMY', false)?.numFmt).toBe('yyyy-mm-dd');
	});

	it('reads a day and month without a year in the current year, in the column order', () => {
		const year = new Date().getFullYear();
		expect(parseOrderedDate('04/03', 'DMY', false)?.value).toBe(at(year, 3, 4));
		expect(parseOrderedDate('04/03', 'MDY', false)?.value).toBe(at(year, 4, 3));
		expect(parseOrderedDate('04/03', 'YDM', false)?.value).toBe(at(year, 3, 4));
		expect(parseOrderedDate('04/03 10:30', 'DMY', false)?.value).toBeCloseTo(
			at(year, 3, 4, 10, 30),
			10,
		);
		expect(parseOrderedDate('31/02', 'DMY', false)).toBeUndefined();
		expect(parseOrderedDate('4.3', 'DMY', false)).toBeUndefined();
	});

	it('accepts one consistent separator, two-digit years and eight-digit compact dates', () => {
		const march4 = at(2024, 3, 4);
		expect(parseOrderedDate('4-3-2024', 'DMY', false)?.value).toBe(march4);
		expect(parseOrderedDate('04.03.24', 'DMY', false)?.value).toBe(march4);
		expect(parseOrderedDate('4 3 2024', 'DMY', false)?.value).toBe(march4);
		expect(parseOrderedDate('04.03.95', 'DMY', false)?.value).toBe(at(1995, 3, 4));
		expect(parseOrderedDate('20240304', 'YMD', false)?.value).toBe(march4);
		expect(parseOrderedDate('04032024', 'DMY', false)?.value).toBe(march4);
		expect(parseOrderedDate('03042024', 'MDY', false)?.value).toBe(march4);
		expect(parseOrderedDate('04/03-2024', 'DMY', false)).toBeUndefined();
	});

	it('rejects values that are not dates in that order', () => {
		expect(parseOrderedDate('31/02/2024', 'DMY', false)).toBeUndefined();
		expect(parseOrderedDate('2024-13-01', 'YMD', false)).toBeUndefined();
		expect(parseOrderedDate('2024-03-04', 'DMY', false)).toBeUndefined();
		expect(parseOrderedDate('04/03/202', 'DMY', false)).toBeUndefined();
		expect(parseOrderedDate('Date', 'DMY', false)).toBeUndefined();
		expect(parseOrderedDate('', 'DMY', false)).toBeUndefined();
	});

	it('reads a time after a space or T, with seconds, fractions and AM/PM', () => {
		expect(parseOrderedDate('2024-03-04T10:20:30', 'YMD', false)).toEqual({
			value: at(2024, 3, 4, 10, 20, 30),
			numFmt: 'yyyy-mm-dd h:mm:ss',
			zoned: false,
		});
		expect(parseOrderedDate('2024-03-04 10:20', 'YMD', false)?.numFmt).toBe('yyyy-mm-dd h:mm');
		const fraction = parseOrderedDate('2024-03-04T10:20:30.250', 'YMD', false);
		expect(fraction?.value).toBeCloseTo(at(2024, 3, 4, 10, 20, 30) + 0.25 / 86_400, 10);
		// The display keeps the fraction (up to milliseconds), so stamps within a second differ.
		expect(fraction?.numFmt).toBe('yyyy-mm-dd h:mm:ss.000');
		expect(parseOrderedDate('2024-03-04T10:20:30.5', 'YMD', false)?.numFmt).toBe(
			'yyyy-mm-dd h:mm:ss.0',
		);
		expect(parseOrderedDate('2024-03-04T10:20:30.123456', 'YMD', false)?.numFmt).toBe(
			'yyyy-mm-dd h:mm:ss.000',
		);
		expect(parseOrderedDate('04/03/2024 2:15 PM', 'DMY', false)).toEqual({
			value: at(2024, 3, 4, 14, 15),
			numFmt: 'yyyy-mm-dd h:mm AM/PM',
			zoned: false,
		});
		expect(parseOrderedDate('2024-03-04T24:00', 'YMD', false)).toBeUndefined();
		expect(parseOrderedDate('2024-03-04T10:61', 'YMD', false)).toBeUndefined();
		// Single-digit minutes and seconds, as typing accepts them, keep the column's order.
		expect(parseOrderedDate('04/03/2024 9:5', 'DMY', false)?.value).toBeCloseTo(
			at(2024, 3, 4, 9, 5),
			10,
		);
		expect(parseOrderedDate('04/03/2024 9:05:7', 'DMY', false)?.value).toBeCloseTo(
			at(2024, 3, 4, 9, 5, 7),
			10,
		);
		// A fraction of a second does not let the hour pass 23.
		expect(parseOrderedDate('2024-03-04T25:00:00.5', 'YMD', false)).toBeUndefined();
		expect(parseOrderedDate('2024-03-04T24:00:00.5', 'YMD', false)).toBeUndefined();
	});

	it('converts a UTC designator or offset to UTC and flags it', () => {
		expect(parseOrderedDate('2024-03-04T10:20:30Z', 'YMD', false)).toEqual({
			value: at(2024, 3, 4, 10, 20, 30),
			numFmt: 'yyyy-mm-dd h:mm:ss',
			zoned: true,
		});
		expect(parseOrderedDate('2024-03-04T10:20:30+02:00', 'YMD', false)?.value).toBeCloseTo(
			at(2024, 3, 4, 8, 20, 30),
			10,
		);
		expect(parseOrderedDate('2024-03-04T10:20:30-0530', 'YMD', false)?.value).toBeCloseTo(
			at(2024, 3, 4, 15, 50, 30),
			10,
		);
		expect(parseOrderedDate('2024-03-04 10:20:30 +01', 'YMD', false)?.value).toBeCloseTo(
			at(2024, 3, 4, 9, 20, 30),
			10,
		);
		expect(parseOrderedDate('2024-03-04T01:00:00+02:00', 'YMD', false)?.value).toBeCloseTo(
			at(2024, 3, 3, 23),
			10,
		);
		expect(parseOrderedDate('2024-03-04T10:20+25:00', 'YMD', false)).toBeUndefined();
	});

	it('reads at most 64 characters, so padding in a crafted file cannot make it slow', () => {
		// A valid shape, but longer than any date and time needs: refused before any pattern runs.
		expect(parseOrderedDate(`2024-03-04${' '.repeat(60)}10:30`, 'YMD', false)).toBeUndefined();
		expect(parseOrderedDate(`2024-03-04${' '.repeat(30_000)}x`, 'YMD', false)).toBeUndefined();
		expect(parseOrderedDate('2024-03-04    10:30', 'YMD', false)?.value).toBeCloseTo(
			at(2024, 3, 4, 10, 30),
			10,
		);
	});

	it('reads fractions of a second with the configured decimal separator', () => {
		const comma = parseOrderedDate('04/03/2024 10:20:30,5', 'DMY', false, ',');
		expect(comma?.value).toBeCloseTo(at(2024, 3, 4, 10, 20, 30) + 0.5 / 86_400, 10);
		expect(comma?.numFmt).toBe('yyyy-mm-dd h:mm:ss.0');
		// With a comma decimal, a period is not a fraction separator.
		expect(parseOrderedDate('04/03/2024 10:20:30.5', 'DMY', false, ',')).toBeUndefined();
	});

	it('honours the 1904 date system', () => {
		expect(parseOrderedDate('2024-03-04', 'YMD', true)?.value).toBe(at(2024, 3, 4) - 1462);
	});
});
