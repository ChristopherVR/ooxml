import { describe, expect, it } from 'vitest';
import { stepDecimals } from './decimals.js';

describe('stepDecimals', () => {
	it.each([
		['General', 1, 1.5, '0.00'],
		['General', -1, 1.5, '0'],
		['General', 1, 5, '0.0'],
		['General', -1, 5, '0'],
		['General', 1, 1 / 3, '0.0000000000'],
		['General', 1, 123456789012, '0.000000E+00'],
		['General', 1, 'text', '0.0'],
		['', 1, 2.25, '0.000'],
	] as const)('%s %d with %s -> %s', (format, delta, sample, expected) => {
		expect(stepDecimals(format, delta, sample)).toBe(expected);
	});

	it.each([
		['0.00', -1, '0.0'],
		['0.0', -1, '0'],
		['0', -1, '0'],
		['0', 1, '0.0'],
		['0.00', 1, '0.000'],
		['#,##0', 1, '#,##0.0'],
		['#,##0.00', -1, '#,##0.0'],
		['0%', 1, '0.0%'],
		['0.0%', -1, '0%'],
		['"$"#,##0.00', 1, '"$"#,##0.000'],
		['[$EUR] #,##0', 1, '[$EUR] #,##0.0'],
		['0.00E+00', -1, '0.0E+00'],
		['0E+00', 1, '0.0E+00'],
		['#,##0.00_);[Red](#,##0.00)', 1, '#,##0.000_);[Red](#,##0.000)'],
		['#,##0_);(#,##0);"-"', 1, '#,##0.0_);(#,##0.0);"-"'],
		['[>100]0.0;0', -1, '[>100]0;0'],
		['0.00 "kg"', -1, '0.0 "kg"'],
	] as const)('%s %d -> %s', (format, delta, expected) => {
		expect(stepDecimals(format, delta)).toBe(expected);
	});

	it('leaves text, dates, times and fractions alone', () => {
		for (const format of ['@', 'm/d/yyyy', 'h:mm:ss', '# ?/?', '"Total"'])
			expect(stepDecimals(format, 1, 3)).toBe(format);
	});
});
