import { describe, expect, it } from 'vitest';
import { BUILTIN_NUMBER_FORMATS, builtinFormatId, FORMAT_PRESETS, formatValue } from './index.js';

describe('built-in number formats', () => {
	it('round-trips every id through builtinFormatId', () => {
		for (const [id, code] of Object.entries(BUILTIN_NUMBER_FORMATS)) {
			expect(builtinFormatId(code)).toBe(Number(id));
		}
	});

	it('knows the ECMA-376 spellings and unknown codes', () => {
		expect(builtinFormatId('mm-dd-yy')).toBe(14);
		expect(builtinFormatId('general')).toBe(0);
		expect(builtinFormatId('m/d/yy h:mm')).toBe(22);
		expect(builtinFormatId('0.000')).toBeUndefined();
	});

	it.each([
		[1, 1234.5, '1235'],
		[3, 1234.5, '1,235'],
		[5, -3.7, '($4)'],
		[8, -1234.5, '($1,234.50)'],
		[10, 0.125, '12.50%'],
		[12, 1.25, '1 1/4'],
		[13, 0.5625, '  9/16'],
		[14, 45000, '3/15/2023'],
		[15, 45000, '15-Mar-23'],
		[18, 45000.75, '6:00 PM'],
		[22, 45000.75, '3/15/2023 18:00'],
		[37, 1234, '1,234 '],
		[38, -1234, '(1,234)'],
		[45, 0.0104166667, '15:00'],
		[46, 1.5, '36:00:00'],
		[47, 0.0104166667, '15:00.0'],
		[48, 1234567, '1.2E+6'],
	])('id %d formats %d like Excel', (id, value, expected) => {
		expect(formatValue(value, BUILTIN_NUMBER_FORMATS[id] ?? '').text).toBe(expected);
	});

	it('lists the ribbon presets', () => {
		expect(FORMAT_PRESETS.map((p) => p.label)).toEqual([
			'General',
			'Number',
			'Currency',
			'Accounting',
			'Short Date',
			'Long Date',
			'Time',
			'Percentage',
			'Fraction',
			'Scientific',
			'Text',
		]);
		expect(formatValue(45000, FORMAT_PRESETS[5]?.format ?? '').text).toBe(
			'Wednesday, March 15, 2023',
		);
	});
});
