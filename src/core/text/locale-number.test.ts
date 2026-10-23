import { describe, expect, it } from 'vitest';
import { formatLocaleNumber, parseLocaleNumber } from './locale-number';

describe('locale numbers', () => {
	it('reads the decimal and group separators of the locale', () => {
		expect(parseLocaleNumber('1,234.5', 'en-US')).toBe(1234.5);
		expect(parseLocaleNumber('1.234,5', 'de-DE')).toBe(1234.5);
		expect(parseLocaleNumber('1,5', 'de-DE')).toBe(1.5);
		expect(parseLocaleNumber('1 234,5', 'fr-FR')).toBe(1234.5);
		expect(parseLocaleNumber(' -2e3 ', 'en-US')).toBe(-2000);
		expect(parseLocaleNumber('.5')).toBe(0.5);
	});

	it('reads a lone point as a decimal in a comma locale, the English paste case', () => {
		expect(parseLocaleNumber('1.5', 'de-DE')).toBe(1.5);
	});

	it('refuses text that is not a number', () => {
		for (const raw of ['', '  ', 'abc', '1,5', '1.2.3', '12a', '1,23']) {
			expect(parseLocaleNumber(raw, 'en-US')).toBeNull();
		}
		expect(parseLocaleNumber('1.2,5', 'de-DE')).toBeNull();
	});

	it('refuses a long run of digits followed by junk without backtracking', () => {
		const start = performance.now();
		expect(parseLocaleNumber(`${'0'.repeat(50_000)}x`, 'en-US')).toBeNull();
		expect(performance.now() - start).toBeLessThan(1000);
	});

	it('writes a value that reads back the same', () => {
		expect(formatLocaleNumber(1234.5, 'de-DE')).toBe('1234,5');
		expect(formatLocaleNumber(1234.5, 'en-US')).toBe('1234.5');
		expect(formatLocaleNumber(Number.NaN)).toBe('');
		for (const value of [0, -3.25, 1e21, 0.1 + 0.2]) {
			expect(parseLocaleNumber(formatLocaleNumber(value, 'de-DE'), 'de-DE')).toBe(value);
		}
	});
});
