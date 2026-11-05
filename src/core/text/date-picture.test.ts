import { describe, expect, it } from 'vitest';

import {
	OOXML_DATETIME_PATTERNS,
	formatDatePicture,
	ooxmlDateTimeFieldPicture,
} from './date-picture';

const date = new Date(2026, 10, 6, 15, 7, 9); // Friday 6 November 2026, 15:07:09 local

describe('formatDatePicture', () => {
	it('formats the numeric date and time tokens', () => {
		expect(formatDatePicture(date, 'M/d/yyyy')).toBe('11/6/2026');
		expect(formatDatePicture(date, 'dd-MM-yy HH:mm:ss')).toBe('06-11-26 15:07:09');
		expect(formatDatePicture(date, 'h:m:s')).toBe('3:7:9');
	});

	it('keeps M (month) and m (minute) apart', () => {
		expect(formatDatePicture(date, 'MM mm')).toBe('11 07');
	});

	it('names months and weekdays in the requested locale', () => {
		expect(formatDatePicture(date, 'dddd, MMMM d, yyyy')).toBe('Friday, November 6, 2026');
		expect(formatDatePicture(date, 'EEE MMM')).toBe('Fri Nov');
		expect(formatDatePicture(date, 'd MMMM yyyy', { locale: 'de-DE' })).toBe('6 November 2026');
		expect(formatDatePicture(date, 'MMMM', { locale: 'fr-FR' })).toBe('novembre');
	});

	it('renders every AM/PM spelling', () => {
		expect(formatDatePicture(date, 'h tt|AM/PM|am/pm|A/P|a/p')).toBe('3 PM|PM|pm|P|p');
		expect(formatDatePicture(new Date(2026, 0, 1, 0, 5), 'h:mm tt')).toBe('12:05 AM');
	});

	it('copies quoted and unknown text literally', () => {
		expect(formatDatePicture(date, "'Week of' d MMM")).toBe('Week of 6 Nov');
		expect(formatDatePicture(date, 'yyyy "year"')).toBe('2026 year');
		expect(formatDatePicture(date, 'yyyy / [x]')).toBe('2026 / [x]');
	});

	it('reads UTC fields when asked', () => {
		const utc = new Date(Date.UTC(2026, 0, 31, 23, 30));
		expect(formatDatePicture(utc, 'yyyy-MM-dd HH:mm', { utc: true })).toBe('2026-01-31 23:30');
		expect(formatDatePicture(utc, 'dddd', { utc: true })).toBe('Saturday');
	});

	it('falls back to English names for an unusable locale', () => {
		expect(formatDatePicture(date, 'MMMM', { locale: 'not a locale!' })).toBe('November');
	});
});

describe('ooxmlDateTimeFieldPicture', () => {
	it('maps every predefined type, ignoring case', () => {
		for (const [type, pattern] of Object.entries(OOXML_DATETIME_PATTERNS)) {
			expect(ooxmlDateTimeFieldPicture(type)).toBe(pattern);
		}
		expect(ooxmlDateTimeFieldPicture('DateTime2')).toBe(OOXML_DATETIME_PATTERNS.datetime2);
	});

	it('decodes the picture a format-encoded type carries', () => {
		expect(ooxmlDateTimeFieldPicture('datetimeyyyy')).toBe('yyyy');
		expect(ooxmlDateTimeFieldPicture('datetimeMMMMyyyy')).toBe('MMMMyyyy');
		expect(ooxmlDateTimeFieldPicture('datetimeyyyy-MM-dd')).toBe('yyyy-MM-dd');
	});

	it('rejects types that name no picture', () => {
		expect(ooxmlDateTimeFieldPicture('datetime')).toBeUndefined();
		expect(ooxmlDateTimeFieldPicture('datetimeFigureOut')).toBeUndefined();
		expect(ooxmlDateTimeFieldPicture('slidenum')).toBeUndefined();
	});
});
