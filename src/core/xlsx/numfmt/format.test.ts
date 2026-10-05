import { describe, expect, it } from 'vitest';
import { cellError } from '../model.js';
import { BUILTIN_NUMBER_FORMATS, formatValue, isDateFormat, OVERFLOW_TEXT } from './index.js';

const fmt = (value: number | string, format: string): string => formatValue(value, format).text;

// Expected strings below were checked against Excel 16 (Range.Text, en-US).
describe('formatValue: numbers', () => {
	it.each([
		[1234.5, '#,##0.00', '1,234.50'],
		[-1234.5, '#,##0.00', '-1,234.50'],
		[1234567, '#,##0,', '1,235'],
		[1234567, '0.0,,"M"', '1.2M'],
		[0.125, '0%', '13%'],
		[-0.125, '0.00%', '-12.50%'],
		[12345.678, '0.00E+00', '1.23E+04'],
		[12345.678, '##0.0E+0', '12.3E+3'],
		[0.000123, '0.00E-00', '1.23E-04'],
		[0.25, '?.??', ' .25'],
		[12.5, '???.???', ' 12.5  '],
		[0, '#.##', '.'],
		[7, '00-00', '00-07'],
		[5551234, '[<=9999999]###-####;(###) ###-####', '555-1234'],
		[123.456, '0.0.0', '123.4.6'],
		[1.005, '0.00', '1.01'],
		[2.5, '0', '3'],
		[-2.5, '0', '-3'],
		[-0.04, '0.0', '0.0'],
		[-0.04, '0.0;-0.0', '-0.0'],
		[1e21, '#,##0', '1,000,000,000,000,000,000,000'],
		[-5, '"x"0', '-x5'],
		[-5, '(0)', '-(5)'],
		[12, '0 \\k\\g', '12 kg'],
		[1234.5, '[$EUR] #,##0.00', 'EUR 1,234.50'],
		[1234.5, '[$€-2] #,##0.00', '€ 1,234.50'],
	])('%d with %s', (value, format, expected) => {
		expect(fmt(value, format)).toBe(expected);
	});
});

describe('formatValue: sections and conditions', () => {
	it.each([
		[-1234.5, '#,##0.00;(#,##0.00)', '(1,234.50)'],
		[0, '#,##0.00;(#,##0.00);"-"', '-'],
		[-1234.5, '0;', ''],
		[0, '0;-0;;@', ''],
		[150, '[>100]"big";[<=100]"small"', 'big'],
		[-5, '[<0]"neg" 0;0', 'neg 5'],
		[-5, '[>100]0.0;0', '-5'],
		[-3, '[>=-5]0;0', '-3'],
		[-3, '[<=-5]0;0', '3'],
		[50, '[<10]0.0', '50'],
		[-10, '[>5]"a"0;"b"0;"c"0', 'b10'],
		[-5, '[>5]"a"0;[<-5]"b"0;"c"0', '-c5'],
		[0, '[>0]0;[<0]0', OVERFLOW_TEXT],
	])('%d with %s', (value, format, expected) => {
		expect(fmt(value, format)).toBe(expected);
	});

	it('reports section colours', () => {
		expect(formatValue(-1234.5, '#,##0.00;[Red](#,##0.00)')).toEqual({
			text: '(1,234.50)',
			color: '#FF0000',
		});
		expect(formatValue(1234.5, '#,##0.00;[Red](#,##0.00)')).toEqual({ text: '1,234.50' });
		expect(formatValue(1234.5, '[Blue]General').color).toBe('#0000FF');
		expect(formatValue(3, '[Color10]0').color).toBe('#008000');
		expect(formatValue(5, '[Green]0').color).toBe('#00FF00');
		expect(formatValue('hi', '[Red]@')).toEqual({ text: 'hi', color: '#FF0000' });
		expect(formatValue(-5, '[Red]0;[Blue]0')).toEqual({ text: '5', color: '#0000FF' });
	});
});

describe('formatValue: text, booleans, errors', () => {
	it('uses the text section for strings', () => {
		expect(fmt('hello', '"<"@">"')).toBe('<hello>');
		expect(fmt('hello', '0;0;0;@" world"')).toBe('hello world');
		expect(fmt('hello', '0.00')).toBe('hello');
		expect(fmt('hi', '0;0;0;"x"')).toBe('x');
		expect(fmt('hi', '"a"@"b"@')).toBe('ahibhi');
		expect(fmt('hi', '@_)')).toBe('hi ');
	});

	it('formats a number with a text-only format as General', () => {
		expect(fmt(1234.5, '@')).toBe('1234.5');
	});

	it('shows booleans, errors and blanks', () => {
		expect(formatValue(true, '0.00').text).toBe('TRUE');
		expect(formatValue(false, 'General').text).toBe('FALSE');
		expect(formatValue(cellError('#DIV/0!'), '0.00').text).toBe('#DIV/0!');
		expect(formatValue(null, '0.00').text).toBe('');
		expect(formatValue(Number.NaN, '0.00').text).toBe('#NUM!');
	});

	it('treats an empty format as General', () => {
		expect(fmt(0.5, '')).toBe('0.5');
	});
});

describe('formatValue: padding and fill', () => {
	// Excel repeats the `*` fill character to the column width; formatValue drops it.
	it.each([
		[1234.5, 44, ' $1,234.50 '],
		[-1234.5, 44, ' $(1,234.50)'],
		[0, 44, ' $-   '],
		[0, 41, ' - '],
		[-5, 41, ' (5)'],
		[1234.5678, -1, '1235'],
	])('%d with builtin %d', (value, id, expected) => {
		const format = id < 0 ? '0*-' : (BUILTIN_NUMBER_FORMATS[id] ?? '');
		expect(fmt(value, format)).toBe(expected);
	});

	it('pads text in accounting formats', () => {
		expect(fmt('abc', BUILTIN_NUMBER_FORMATS[42] ?? '')).toBe(' abc ');
	});
});

describe('formatValue: dates and times', () => {
	it.each([
		[45000, 'm/d/yyyy', '3/15/2023'],
		[45000, 'dddd, mmmm d, yyyy', 'Wednesday, March 15, 2023'],
		[45000.75, 'h:mm AM/PM', '6:00 PM'],
		[45000.999999, 'h:mm:ss', '0:00:00'],
		[45000.9999999, 'yyyy-mm-dd hh:mm:ss', '2023-03-16 00:00:00'],
		[2.75, '[h]:mm', '66:00'],
		[0.0625, '[mm]:ss', '90:00'],
		[0.123456, 'h:mm:ss.00', '2:57:46.60'],
		[45000.5125, 'yy mm ss', '23 18 00'],
		[60, 'm/d/yyyy', '2/29/1900'],
		[0, 'm/d/yyyy', '1/0/1900'],
		[-0.5, '[h]:mm', OVERFLOW_TEXT],
		[2958466, 'm/d/yyyy', OVERFLOW_TEXT],
		// A lone elapsed unit is a signed count with no date limit (Excel 16).
		[-0.5, '[s]', '-43200'],
		[-0.5, '[h] "h"', '-12 h'],
		[1e15, '[h]', '24000000000000000'],
		[-0.5, '[s].00', OVERFLOW_TEXT],
		[45000, '[$-F800]dddd, mmmm dd, yyyy', 'Wednesday, March 15, 2023'],
		[45000, 'ggge"年"m"月"d"日"', '2023年3月15日'],
		[45000, 'bbbb', '2566'],
	])('%d with %s', (value, format, expected) => {
		expect(fmt(value, format)).toBe(expected);
	});

	it('supports the 1904 date system', () => {
		expect(formatValue(0, 'm/d/yyyy', { date1904: true }).text).toBe('1/1/1904');
		expect(formatValue(45000, 'dddd', { date1904: true }).text).toBe('Tuesday');
	});
});

// Recorded from Excel 16: the last continued-fraction convergent that fits, never a semiconvergent.
describe('formatValue: fractions', () => {
	it.each([
		[0.06, '?/?', '0/1'],
		[0.3, '?/?', '2/7'],
		[0.09, '??/??', ' 1/11'],
		[0.13, '??/??', ' 3/23'],
		[84152.9034, '?/?', '84153/1'],
		[1000.11, '# ??/??', '1000  1/9 '],
		[0.6, '# ?/8', ' 5/8'],
		[0.5, '# #/#', '1/2'],
		[1, '# #/#', '1'],
		[0, '# #/# "x"', '0 x'],
		[0.5, '# #/?', '1/2'],
		[0.5, '# ?/#', ' 1/2'],
		[-0.0001, '# ?/?', '-0    '],
		[-0.0001, '?/?', '0/1'],
		[0, '##0.0E+0', '000.0E+0'],
	])('%d with %s', (value, format, expected) => {
		expect(fmt(value, format)).toBe(expected);
	});
});

describe('isDateFormat', () => {
	it.each([
		['m/d/yyyy', true],
		['[h]:mm:ss', true],
		['h:mm AM/PM', true],
		['[$-409]mmmm d, yyyy;@', true],
		['General', false],
		['0.00E+00', false],
		['#,##0.00;[Red](#,##0.00)', false],
		['0.00"d"', false],
		['0 \\d', false],
		['@', false],
		['[Red][>100]0', false],
	])('%s -> %s', (format, expected) => {
		expect(isDateFormat(format)).toBe(expected);
	});
});
