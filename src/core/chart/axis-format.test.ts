import { describe, expect, it } from 'vitest';

import { formatCompactAxisValue } from './axis-format';

describe('formatCompactAxisValue', () => {
	it('abbreviates thousands and millions', () => {
		expect(formatCompactAxisValue(1500)).toBe('1.5K');
		expect(formatCompactAxisValue(-2_500_000)).toBe('-2.5M');
	});

	it('keeps integers and rounds fractions to one decimal', () => {
		expect(formatCompactAxisValue(42)).toBe('42');
		expect(formatCompactAxisValue(0.25)).toBe('0.3');
	});

	it('lets an explicit format code win over the compact form', () => {
		expect(formatCompactAxisValue(0.5, '0%')).toBe('50%');
		expect(formatCompactAxisValue(1500, '#,##0')).toBe('1,500');
		expect(formatCompactAxisValue(1500, 'General')).toBe('1.5K');
	});
});
