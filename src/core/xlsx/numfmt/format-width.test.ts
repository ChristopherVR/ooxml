import { describe, expect, it } from 'vitest';
import { formatValue } from './format.js';
import { formatGeneral } from './general.js';

describe('formatGeneral width', () => {
	it.each([
		[3.14159265, 4, '3.14'],
		[3.14159265, 1, '3'],
		[123456, 5, '1E+05'],
		[123456, 6, '123456'],
		[1234567.891, 8, '1234568'],
		[0.000123456, 6, '0.0001'],
		[1 / 3, 5, '0.333'],
	])('%d in %d chars -> %s', (value, width, expected) => {
		expect(formatGeneral(value, width)).toBe(expected);
	});

	it('keeps the 11-character default', () => {
		expect(formatGeneral(1 / 3)).toBe('0.333333333');
	});
});

describe('formatValue with width', () => {
	it('is unchanged without a width', () => {
		expect(formatValue(1 / 3, 'General').text).toBe('0.333333333');
		expect(formatValue(1234567, '#,##0').text).toBe('1,234,567');
	});

	it('shortens General numbers to the available characters', () => {
		expect(formatValue(1 / 3, 'General', { width: 6 }).text).toBe('0.3333');
		expect(formatValue(-1 / 3, 'General', { width: 6 }).text).toBe('-0.333');
		expect(formatValue(123456789, 'General', { width: 7 }).text).toBe('1.2E+08');
		expect(formatValue(42, 'General', { width: 11 }).text).toBe('42');
	});

	it('fills non-General numbers and dates that do not fit with #', () => {
		expect(formatValue(1234567, '#,##0', { width: 5 }).text).toBe('#####');
		expect(formatValue(45000, 'mmmm d, yyyy', { width: 4 }).text).toBe('####');
		expect(formatValue(123456, 'General', { width: 3 }).text).toBe('###');
	});

	it('uses a measure function with a pixel width', () => {
		const measure = (text: string): number => text.length * 7;
		expect(formatValue(1 / 3, 'General', { width: 42, measure }).text).toBe('0.3333');
		expect(formatValue(1234567, '0', { width: 35, measure }).text).toBe('#####');
	});

	it('never fits text, booleans or errors', () => {
		expect(formatValue('a long text value', 'General', { width: 2 }).text).toBe(
			'a long text value',
		);
		expect(formatValue(true, 'General', { width: 1 }).text).toBe('TRUE');
	});

	it('keeps the section colour on a hash fill', () => {
		expect(formatValue(-123456, '[Red]0', { width: 3 })).toEqual({ text: '###', color: '#FF0000' });
	});
});
