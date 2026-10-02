import { describe, expect, it } from 'vitest';
import { serialToYmd, weekdayOf, ymdToSerial } from './date-serial.js';
import { numberToText, parseNumberText } from './text-number.js';

// Expected values were read from Excel 16 (`=""&x` and `="text"+0`).
describe('numberToText', () => {
	it.each([
		[1 / 3, '0.333333333333333'],
		[2 / 3, '0.666666666666667'],
		[1e15, '1000000000000000'],
		[1e19, '10000000000000000000'],
		[1e20, '1E+20'],
		[1.23456789e25, '1.23456789E+25'],
		[0.00001, '0.00001'],
		[1e-10, '0.0000000001'],
		[1e-18, '0.000000000000000001'],
		[1e-19, '1E-19'],
		[-1.5e-20, '-1.5E-20'],
		[1.42857142857143e-6, '1.42857142857143E-06'],
		[0.1 + 0.2, '0.3'],
		[1234567.891, '1234567.891'],
		[-0, '0'],
		[1e301, '1E+301'],
		[(1 / 7) * 1e17, '14285714285714300'],
	])('%s -> %s', (value, text) => {
		expect(numberToText(value)).toBe(text);
	});
});

describe('parseNumberText', () => {
	it.each([
		['$1,000', 1000],
		['10%', 0.1],
		['50 %', 0.5],
		[' 3 ', 3],
		['(5)', -5],
		['1e3', 1000],
		['1E+3', 1000],
		['5.', 5],
		['.5', 0.5],
		['+5', 5],
		['$-5', -5],
		['-$5', -5],
		['1,234.5', 1234.5],
		['1 1/2', 1.5],
		['2020-01-15', 43845],
		['15-Jan-2020', 43845],
		['Jan 15, 2020', 43845],
		['2020/01/15', 43845],
		['Jan-2020', 43831],
		['12:30', 0.5208333333333334],
		['12:30 PM', 0.5208333333333334],
		['3:00 AM', 0.125],
		['25:00', 1.0416666666666667],
		['2020-1-5 10:30', 43835.4375],
	])('%j -> %s', (text, value) => {
		expect(parseNumberText(text)).toBeCloseTo(value, 10);
	});

	it.each(['', '.', '--5', '5-', '1,23', 'TRUE', '5%%', 'abc', '13/45/2020'])(
		'%j is not a number',
		(text) => {
			expect(parseNumberText(text)).toBeUndefined();
		},
	);
});

describe('date serials', () => {
	it('honours the 1900 leap-year bug', () => {
		expect(serialToYmd(60)).toEqual({ year: 1900, month: 2, day: 29 });
		expect(serialToYmd(59)).toEqual({ year: 1900, month: 2, day: 28 });
		expect(serialToYmd(61)).toEqual({ year: 1900, month: 3, day: 1 });
		expect(serialToYmd(0)).toEqual({ year: 1900, month: 1, day: 0 });
		expect(ymdToSerial(1900, 2, 29)).toBe(60);
		expect(ymdToSerial(1900, 3, 1)).toBe(61);
		expect(ymdToSerial(2021, 1, 1)).toBe(44197);
		expect(ymdToSerial(2020, 13, 1)).toBe(44197);
	});

	it('supports the 1904 system', () => {
		expect(ymdToSerial(1904, 1, 1, true)).toBe(0);
		expect(ymdToSerial(2021, 1, 1, true)).toBe(44197 - 1462);
		expect(serialToYmd(0, true)).toEqual({ year: 1904, month: 1, day: 1 });
		expect(weekdayOf(0, true)).toBe(5);
	});

	it('reports serial 1 as a Sunday like Excel', () => {
		expect(weekdayOf(1)).toBe(0);
		expect(weekdayOf(44197)).toBe(5);
	});
});
