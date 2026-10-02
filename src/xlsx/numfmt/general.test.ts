import { describe, expect, it } from 'vitest';
import { formatGeneral } from './general.js';

// Recorded from Excel 16: TEXT/Range.Text of a General cell wide enough for 11 characters.
describe('formatGeneral', () => {
	it.each([
		[0, '0'],
		[0.1 + 0.2, '0.3'],
		[1 / 3, '0.333333333'],
		[-2 / 3, '-0.666666667'],
		[12345678901, '12345678901'],
		[-12345678901, '-12345678901'],
		[123456789012, '1.23457E+11'],
		[99999999999.5, '1E+11'],
		[1234567890.12345, '1234567890'],
		[0.00001, '0.00001'],
		[1e-9, '0.000000001'],
		[1e-10, '1E-10'],
		[1.23456789e-5, '1.23457E-05'],
		[0.000123456789, '0.000123457'],
		[0.0009999999999, '0.001'],
		[1e100, '1E+100'],
		[-1.5e-300, '-1.5E-300'],
	])('%d -> %s', (value, expected) => {
		expect(formatGeneral(value)).toBe(expected);
	});
});
