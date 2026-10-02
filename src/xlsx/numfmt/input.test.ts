import { describe, expect, it } from 'vitest';
import { cellError } from '../model.js';
import { formatValue, parseCellInput } from './index.js';

// Values and formats recorded by typing each text into a General cell of Excel 16 (en-US).
// Where Excel used the machine's short date (yyyy-mm-dd here), the en-US default m/d/yyyy is
// expected instead; `€` input is formatted (Excel en-US keeps it General).
describe('parseCellInput', () => {
	it.each([
		['1,234', 1234, '#,##0'],
		['1,234.5', 1234.5, '#,##0.00'],
		['1,234,567.891', 1234567.891, '#,##0.00'],
		['$1,234.5', 1234.5, '$#,##0.00_);[Red]($#,##0.00)'],
		['$1234', 1234, '$#,##0_);[Red]($#,##0)'],
		['-$5', -5, '$#,##0_);[Red]($#,##0)'],
		['$-5', -5, '$#,##0_);[Red]($#,##0)'],
		['$ 12', 12, '$#,##0_);[Red]($#,##0)'],
		['£5.25', 5.25, '£#,##0.00'],
		['€5', 5, '€#,##0'],
		['10%', 0.1, '0%'],
		['12.5%', 0.125, '0.00%'],
		['1,234%', 12.34, '0%'],
		['12 %', 0.12, '0%'],
		['1.5E+3', 1500, '0.00E+00'],
		['1e5', 100000, '0.00E+00'],
		['-1.5e-3', -0.0015, '0.00E+00'],
		['(123)', -123, undefined],
		['(1,234.50)', -1234.5, '#,##0.00'],
		['  12  ', 12, undefined],
		['+5', 5, undefined],
		['.5', 0.5, undefined],
		['00123', 123, undefined],
		['-0', 0, undefined],
		['0 1/2', 0.5, '# ?/?'],
		['1 3/4', 1.75, '# ?/?'],
		['2023-03-15', 45000, 'yyyy-mm-dd'],
		['2023/3/15', 45000, 'yyyy-mm-dd'],
		['3/15/2023', 45000, 'm/d/yyyy'],
		['3/15/23', 45000, 'm/d/yyyy'],
		['3-15-2023', 45000, 'm/d/yyyy'],
		['15-Mar-2023', 45000, 'd-mmm-yy'],
		['15 March 2023', 45000, 'd-mmm-yy'],
		['March 15, 2023', 45000, 'd-mmm-yy'],
		['Jan 2023', 44927, 'mmm-yy'],
		['Mar-2023', 44986, 'mmm-yy'],
		['9999-12-31', 2958465, 'yyyy-mm-dd'],
		['1900-02-29', 60, 'yyyy-mm-dd'],
		['13:45', 0.572916666666667, 'h:mm'],
		['1:5', 0.0451388888888889, 'h:mm'],
		['1:45 PM', 0.572916666666667, 'h:mm AM/PM'],
		['12:30 am', 0.0208333333333333, 'h:mm AM/PM'],
		['13:45:30', 0.573263888888889, 'h:mm:ss'],
		['8:30:15 PM', 0.854340277777778, 'h:mm:ss AM/PM'],
		['25:00', 1.04166666666667, '[h]:mm:ss'],
		['45:30.5', 0.0316030092592593, 'mm:ss.0'],
		['3/15/2023 13:45', 45000.5729166667, 'm/d/yyyy h:mm'],
		['2023-03-15 08:30', 45000.3541666667, 'yyyy-mm-dd h:mm'],
	])('%s', (text, value, numFmt) => {
		const parsed = parseCellInput(text);
		expect(parsed.value).toBeCloseTo(value, 9);
		expect(parsed.numFmt).toBe(numFmt);
		expect(parsed.formula).toBeUndefined();
	});

	it.each(['12,34', '1.2.3', '5-', '5e', '1899-12-31', '2023-02-30', 'hello', '13:75', '$1e3'])(
		'keeps %s as text',
		(text) => {
			expect(parseCellInput(text)).toEqual({ value: text });
		},
	);

	it('handles formulas, forced text, booleans, errors and blanks', () => {
		expect(parseCellInput('=SUM(A1:A3)')).toEqual({ value: null, formula: 'SUM(A1:A3)' });
		expect(parseCellInput("'00123")).toEqual({ value: '00123' });
		expect(parseCellInput("'=1+1")).toEqual({ value: '=1+1' });
		expect(parseCellInput('TRUE')).toEqual({ value: true });
		expect(parseCellInput('false')).toEqual({ value: false });
		expect(parseCellInput('#N/A')).toEqual({ value: cellError('#N/A') });
		expect(parseCellInput('#div/0!')).toEqual({ value: cellError('#DIV/0!') });
		expect(parseCellInput('')).toEqual({ value: null });
		expect(parseCellInput('   ')).toEqual({ value: '   ' });
		expect(parseCellInput('=')).toEqual({ value: '=' });
	});

	it('uses the current year for day-month input', () => {
		const parsed = parseCellInput('15-Mar');
		expect(parsed.numFmt).toBe('d-mmm');
		expect(formatValue(parsed.value, 'yyyy').text).toBe(String(new Date().getFullYear()));
	});

	it('produces 1904 serials when asked', () => {
		expect(parseCellInput('1904-01-02', { date1904: true }).value).toBe(1);
		expect(parseCellInput('1903-12-31', { date1904: true }).value).toBe('1903-12-31');
	});

	it('round-trips typed values through their format', () => {
		for (const text of ['1,234.50', '12.50%', '3/15/2023', '1:45 PM', '1 3/4', '$1,234.50 ']) {
			const parsed = parseCellInput(text.trim());
			expect(formatValue(parsed.value, parsed.numFmt ?? 'General').text).toBe(text);
		}
	});
});
