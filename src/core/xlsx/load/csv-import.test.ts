import { describe, expect, it } from 'vitest';
import { parseAddress } from '../address';
import { getCell } from '../cells';
import { dateToSerial } from '../numfmt/date';
import { internStyle, styleAt } from '../styles';
import { csvToWorkbook, sheetToCsv, type CsvWorkbookOptions } from './csv';

/** Loads `text` (comma-delimited unless told otherwise) and reads cells by A1 reference. */
function load(text: string, options: CsvWorkbookOptions = {}) {
	const workbook = csvToWorkbook(text, { delimiter: ',', ...options });
	const sheet = workbook.sheets[0]!;
	const cell = (ref: string) => {
		const at = parseAddress(ref)!;
		return getCell(sheet, at.row, at.col);
	};
	return {
		workbook,
		warnings: workbook.warnings,
		value: (ref: string) => cell(ref)?.value,
		format: (ref: string) => styleAt(workbook, cell(ref)?.styleId).numFmt,
	};
}

const day = (year: number, month: number, date: number): number =>
	dateToSerial(new Date(Date.UTC(year, month - 1, date)));

describe('csvToWorkbook default conversions', () => {
	it('types fields exactly as before', () => {
		const csv = load(
			'00123,12345678901234567890,04/03/2024,"1,5",2024-03-04T10:20:30Z,123E5,JAN1\n',
		);
		expect(csv.value('A1')).toBe(123);
		// The double nearest the 20 digits (Excel would keep the first 15: 12345678901234500000).
		expect(csv.value('B1')).toBe(Number('12345678901234567890'));
		expect(csv.value('C1')).toBe(day(2024, 4, 3));
		expect(csv.value('D1')).toBe('1,5');
		expect(csv.value('E1')).toBe('2024-03-04T10:20:30Z');
		expect(csv.value('F1')).toBe(12_300_000);
		expect(typeof csv.value('G1')).toBe('number');
	});

	it('reports each lossy conversion once per column and kind, with a count and the first cell', () => {
		const csv = load(
			'id,acct,code,gene\n007,1234567890123456,1E5,JAN1\n0042,12345678901234567890,2E3,SEPT2\n',
		);
		expect(csv.warnings).toEqual([
			'Column A: 2 values lost their leading zeros when converted to numbers (first: A2).',
			'Column B: 1 value with more than 15 digits lost digits when converted to a number (first: B3).',
			'Column D: 2 values of letters and digits such as JAN1 were converted to dates (first: D2).',
		]);
	});

	it('words a single conversion in the singular and stays quiet when nothing is lost', () => {
		expect(load('a\n0123\n').warnings).toEqual([
			'Column A: 1 value lost its leading zeros when converted to a number (first: A2).',
		]);
		expect(
			load('Name,Age,Share,When\nAda,36,0.5,2024-01-15\nGrace,045.5%,1E,Jan 5\n').warnings,
		).toEqual(['Column B: 1 value lost its leading zeros when converted to a number (first: B3).']);
		expect(load('n\n0\n0.25\n-0.5\n100\n000000000000000123456\n').warnings).toEqual([
			'Column A: 1 value lost its leading zeros when converted to a number (first: A6).',
		]);
		// Zeros after the decimal separator are not leading zeros.
		expect(load('.05,-.05,$.05,5%\n').warnings).toEqual([]);
		expect(load('",05";"-,007"\n', { delimiter: ';', decimalSeparator: ',' }).warnings).toEqual([]);
	});

	it('adds no warnings when reportConversions is off', () => {
		const csv = load('007,1E5\n', { reportConversions: false });
		expect(csv.value('A1')).toBe(7);
		expect(csv.warnings).toEqual([]);
	});
});

describe('csvToWorkbook conversion switches', () => {
	it('keeps leading zeros as text when leadingZeros is off', () => {
		const csv = load('00123,0.5,123,-0012,$00123,007.5%,00,.05\n', {
			conversions: { leadingZeros: false },
		});
		expect(['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1'].map(csv.value)).toEqual([
			'00123',
			0.5,
			123,
			'-0012',
			'$00123',
			'007.5%',
			'00',
			0.05,
		]);
		expect(csv.format('A1')).toBe('General');
		expect(csv.warnings).toEqual([]);
		// A zero before a thousands separator leads the integer part too.
		expect(load('"0,123"\n', { conversions: { leadingZeros: false } }).value('A1')).toBe('0,123');
		expect(load('"0,123"\n').warnings).toEqual([
			'Column A: 1 value lost its leading zeros when converted to a number (first: A1).',
		]);
		const comma = load('0.123;0,5\n', {
			delimiter: ';',
			decimalSeparator: ',',
			conversions: { leadingZeros: false },
		});
		expect(['A1', 'B1'].map(comma.value)).toEqual(['0.123', 0.5]);
	});

	it('keeps numbers of 16 or more digits as text when longNumbers is off', () => {
		const csv = load(
			'12345678901234567890,123456789012345,1234567890123456,0000000000000000123,1234567890.123456\n',
			{ conversions: { longNumbers: false } },
		);
		expect(csv.value('A1')).toBe('12345678901234567890');
		expect(csv.value('B1')).toBe(123456789012345);
		expect(csv.value('C1')).toBe('1234567890123456');
		// Leading zeros are not digits of precision.
		expect(csv.value('D1')).toBe(123);
		expect(csv.value('E1')).toBe('1234567890.123456');
	});

	it('keeps digits around the letter E as text when scientific is off', () => {
		const csv = load('123E5,1.5e-3,1500,E5\n', { conversions: { scientific: false } });
		expect(['A1', 'B1', 'C1', 'D1'].map(csv.value)).toEqual(['123E5', '1.5e-3', 1500, 'E5']);
	});

	it('keeps continuous letters and digits as text when letterDates is off', () => {
		const csv = load('JAN1,MARCH1,sept2,Jan 1,Jan-1,10:30am, JAN1\n', {
			conversions: { letterDates: false },
		});
		expect(['A1', 'B1', 'C1'].map(csv.value)).toEqual(['JAN1', 'MARCH1', 'sept2']);
		// Padding is trimmed before typing, so it does not hide the letters from the switch.
		expect(csv.value('G1')).toBe(' JAN1');
		expect(load('id, JAN1\n').warnings).toEqual([
			'Column B: 1 value of letters and digits such as JAN1 was converted to a date (first: B1).',
		]);
		// Excel documents that values with a space or another character may still become dates.
		expect(typeof csv.value('D1')).toBe('number');
		expect(typeof csv.value('E1')).toBe('number');
		expect(csv.value('F1')).toBeCloseTo(10.5 / 24);
	});

	it('reports long numbers only when the stored number does not keep their digits', () => {
		const csv = load(
			'a,b,c,d\n0.30000000000000004,1234567890123456,1.0000000000000000,12345678901234567890\n',
		);
		expect(csv.value('A2')).toBe(0.30000000000000004);
		expect(csv.warnings).toEqual([
			'Column D: 1 value with more than 15 digits lost digits when converted to a number (first: D2).',
		]);
		// The switch still follows Excel's rule: 16 or more digits stay text when it is off.
		const off = load('0.30000000000000004\n', { conversions: { longNumbers: false } });
		expect(off.value('A1')).toBe('0.30000000000000004');
	});

	it('reports scientific notation only when the number does not keep what the text says', () => {
		// Our own export writes 1.00E+05; reading it back changes nothing.
		const workbook = csvToWorkbook('x\n', { delimiter: ',' });
		const sheet = workbook.sheets[0]!;
		const scientific = internStyle(workbook, { ...styleAt(workbook, 0), numFmt: '0.00E+00' });
		sheet.rows.set(0, new Map([[0, { value: 100_000, styleId: scientific }]]));
		const back = load(sheetToCsv(workbook, 0));
		expect(back.value('A1')).toBe(100_000);
		expect(back.warnings).toEqual([]);
		expect(load('123E5,1.5e-10\n').warnings).toEqual([]);
		// An exponent below what a number holds turns the value into 0.
		expect(load('1.5e-400\n').warnings).toEqual([
			'Column A: 1 value with the letter E lost digits when converted to a number (first: A1).',
		]);
	});

	it('keeps a field as text when any conversion it needs is off', () => {
		const csv = load('00012345678901234567\n', { conversions: { longNumbers: false } });
		expect(csv.value('A1')).toBe('00012345678901234567');
	});
});

describe('csvToWorkbook column formats', () => {
	it('imports text columns verbatim with the Text number format', () => {
		const csv = load('00123,=1+2,TRUE,1E5\n', {
			formulas: true,
			columns: { 0: 'text', 1: 'text', 2: 'text', 3: 'general' },
		});
		expect(['A1', 'B1', 'C1'].map(csv.value)).toEqual(['00123', '=1+2', 'TRUE']);
		expect(csv.format('A1')).toBe('@');
		expect(csv.value('D1')).toBe(100_000);
		// 1E5 is exactly the number its text says, so nothing is reported.
		expect(csv.warnings).toEqual([]);
	});

	it('skips columns and moves the following ones left', () => {
		const csv = load('a,secret,b,c,d\n1,x,0002,3,4\n', { columns: { 1: 'skip', 3: 'skip' } });
		expect(['A1', 'B1', 'C1', 'D1'].map(csv.value)).toEqual(['a', 'b', 'd', undefined]);
		expect(['A2', 'B2', 'C2'].map(csv.value)).toEqual([1, 2, 4]);
		expect(csv.warnings).toEqual([
			'Column B: 1 value lost its leading zeros when converted to a number (first: B2).',
		]);
		// An array of formats works too.
		expect(load('1,2,3\n', { columns: ['text', 'skip'] }).value('B1')).toBe(3);
	});

	it('reads date columns in the given order and falls back to General otherwise', () => {
		const csv = load('When,Due\n04/03/2024,2024-03-04\n31.12.24,not a date\n', {
			columns: { 0: 'DMY', 1: 'DMY' },
		});
		expect(csv.value('A1')).toBe('When');
		expect(csv.value('A2')).toBe(day(2024, 3, 4));
		expect(csv.format('A2')).toBe('yyyy-mm-dd');
		expect(csv.value('A3')).toBe(day(2024, 12, 31));
		// A year-first date in a DMY column is read as General, as the wizard documents.
		expect(csv.value('B2')).toBe(day(2024, 3, 4));
		expect(csv.value('B3')).toBe('not a date');
	});

	it('reports date-column values that General read as dates in another order', () => {
		const csv = load(
			'When,Note\n04/03/2024,x\n04/13/2024,y\n2024-03-05,z\nn/a,w\n10:30,v\nMar 5,u\n4/13,t\n',
			{ columns: { 0: 'DMY' } },
		);
		// The wizard's General fallback keeps its value; the column now mixes two orders.
		expect(csv.value('A3')).toBe(day(2024, 4, 13));
		// Only month-first numeric reads are reported: not ISO dates, times or month names.
		expect(csv.warnings).toEqual([
			'Column A: 2 values could not be read as dates in DMY order and were read as General (first: A3).',
		]);
		// In a month-first column the General reading agrees, so there is nothing to report.
		expect(load('2024/13/04\n', { columns: { 0: 'YMD' } }).warnings).toEqual([]);
		expect(load('04/13/2024\n', { columns: { 0: 'DMY' } }).warnings).toEqual([
			'Column A: 1 value could not be read as a date in DMY order and was read as General (first: A1).',
		]);
	});

	it('converts zoned date-times to UTC and reports them once per column', () => {
		const csv = load(
			'id,stamp\n1,2024-03-04T10:20:30Z\n2,2024-03-04T12:00:00+02:00\n3,2024-03-04\n',
			{
				columns: { 0: 'skip', 1: 'YMD' },
			},
		);
		expect(csv.value('A2')).toBe(day(2024, 3, 4) + (10 * 3600 + 20 * 60 + 30) / 86_400);
		expect(csv.value('A3')).toBeCloseTo(day(2024, 3, 4) + 10 / 24, 10);
		expect(csv.format('A3')).toBe('yyyy-mm-dd h:mm:ss');
		expect(csv.warnings).toEqual([
			'Column A: 2 date-times with a time zone were converted to UTC (first: A2).',
		]);
		expect(load('2024-03-04T10:20Z\n', { columns: { 0: 'YMD' } }).warnings).toEqual([
			'Column A: 1 date-time with a time zone was converted to UTC (first: A1).',
		]);
	});

	it('rejects unknown column formats and separators that collide', () => {
		expect(() => load('a\n', { columns: { 0: 'number' as 'text' } })).toThrow(RangeError);
		expect(() => load('a\n', { columns: { [-1]: 'text' } })).toThrow(RangeError);
		expect(() => load('a\n', { decimalSeparator: ',', thousandsSeparator: ',' })).toThrow(
			RangeError,
		);
		expect(() => load('a\n', { decimalSeparator: '' })).toThrow(RangeError);
	});
});

describe('csvToWorkbook separators and trailing minus', () => {
	it('reads a comma decimal separator, swapping the thousands separator to a period', () => {
		const csv = load('1,5;1.234,56;12,5%;3.14;3.140;04.03.2024;Smith, J.\n', {
			delimiter: ';',
			decimalSeparator: ',',
		});
		expect(csv.value('A1')).toBe(1.5);
		expect(csv.value('B1')).toBe(1234.56);
		expect(csv.format('B1')).toBe('#,##0.00');
		expect(csv.value('C1')).toBeCloseTo(0.125);
		// A period is not a decimal separator here, so `3.14` stays text; `3.140` is grouped.
		expect(csv.value('D1')).toBe('3.14');
		expect(csv.value('E1')).toBe(3140);
		expect(csv.value('F1')).toBe('04.03.2024');
		expect(csv.value('G1')).toBe('Smith, J.');
		// A time with a comma fraction reads once localized; a plain date keeps its own reading.
		const stamp = load('04/03/2024 10:20:30,5\n', {
			delimiter: ';',
			decimalSeparator: ',',
			columns: { 0: 'DMY' },
		});
		expect(stamp.value('A1')).toBeCloseTo(
			day(2024, 3, 4) + (10 * 3600 + 20 * 60 + 30.5) / 86_400,
			10,
		);
		expect(stamp.warnings).toEqual([]);
		const times = load('10:30:15,5;Mar 5, 2024\n', { delimiter: ';', decimalSeparator: ',' });
		expect(times.value('A1')).toBeCloseTo((10 * 3600 + 30 * 60 + 15.5) / 86_400, 10);
		expect(times.value('B1')).toBe(day(2024, 3, 5));
	});

	it('reads a space as the thousands separator and keeps formulas and times intact', () => {
		const csv = load('1 234 567.5,New York,=SUM(1.5),10:20:30.5\n', {
			thousandsSeparator: ' ',
			formulas: true,
		});
		expect(csv.value('A1')).toBe(1234567.5);
		expect(csv.value('B1')).toBe('New York');
		expect(csv.workbook.sheets[0]!.rows.get(0)?.get(2)?.formula).toBe('SUM(1.5)');
		expect(csv.value('D1')).toBeCloseTo((10 * 3600 + 20 * 60 + 30.5) / 86_400);
		// Only a separator between digits groups thousands; padding and currency spacing stay.
		const euro = load('1 234,56 €; 1 234,5 ;€ 12 345\n', {
			delimiter: ';',
			decimalSeparator: ',',
			thousandsSeparator: ' ',
		});
		expect(['A1', 'B1', 'C1'].map(euro.value)).toEqual([1234.56, 1234.5, 12_345]);
		expect(euro.format('A1')).toBe('€#,##0.00');
		// A fraction's space is not a thousands separator.
		const fraction = load('1 1/2;12 345;2 3/4\n', {
			delimiter: ';',
			decimalSeparator: ',',
			thousandsSeparator: ' ',
		});
		expect(['A1', 'B1', 'C1'].map(fraction.value)).toEqual([1.5, 12_345, 2.75]);
		expect(fraction.format('A1')).toBe('# ?/?');
		// Nor is it a grouped zero: `0 1/2` has no leading zero to lose, `00 1/2` has.
		const half = { delimiter: ';', thousandsSeparator: ' ' } as const;
		expect(load('0 1/2\n', half).warnings).toEqual([]);
		const kept = load('0 1/2;00 1/2\n', { ...half, conversions: { leadingZeros: false } });
		expect(['A1', 'B1'].map(kept.value)).toEqual([0.5, '00 1/2']);
	});

	it('reads a trailing minus as a negative number only when asked', () => {
		const text = '123-;1.234,56-;-5-;abc-;-;00012-;$-5-;€ -5-;+5-;$5-\n';
		const plain = load(text, { delimiter: ';' });
		expect(plain.value('A1')).toBe('123-');
		const csv = load(text, { delimiter: ';', decimalSeparator: ',', trailingMinus: true });
		expect(['A1', 'B1', 'C1', 'D1', 'E1', 'F1'].map(csv.value)).toEqual([
			-123,
			-1234.56,
			'-5-',
			'abc-',
			'-',
			-12,
		]);
		// A value that already has a sign, even after a currency symbol, is not negated again.
		expect(['G1', 'H1', 'I1', 'J1'].map(csv.value)).toEqual(['$-5-', '€ -5-', '+5-', -5]);
		expect(csv.warnings).toEqual([
			'Column F: 1 value lost its leading zeros when converted to a number (first: F1).',
		]);
	});
});

describe('csvToWorkbook with an export from another system', () => {
	it('keeps IDs and account numbers, reads day-first dates and comma decimals', () => {
		const text = [
			'Customer;Account;Opened;Balance;Region',
			'000417;40012345678901234567;04/03/2024;1.234,50;0007',
			'000981;40098765432109876543;31/12/2023;-87,25;0012',
			'001502;40011112222333344445;01/02/2024;0,00;0003',
		].join('\r\n');
		const csv = load(text, {
			delimiter: ';',
			decimalSeparator: ',',
			columns: { 0: 'text', 1: 'text', 2: 'DMY' },
			conversions: { leadingZeros: false },
		});
		expect(csv.value('A2')).toBe('000417');
		expect(csv.value('B3')).toBe('40098765432109876543');
		expect(csv.value('C2')).toBe(day(2024, 3, 4));
		expect(csv.value('C3')).toBe(day(2023, 12, 31));
		expect(csv.value('C4')).toBe(day(2024, 2, 1));
		expect(csv.value('D2')).toBe(1234.5);
		expect(csv.value('D3')).toBe(-87.25);
		expect(csv.value('D4')).toBe(0);
		expect(csv.value('E2')).toBe('0007');
		expect(csv.warnings).toEqual([]);
		// The same file with defaults loses the IDs and the account digits, and says so.
		const naive = load(text, { delimiter: ';' });
		expect(naive.value('A2')).toBe(417);
		expect(naive.value('C2')).toBe(day(2024, 4, 3));
		expect(naive.value('D2')).toBe('1.234,50');
		expect(naive.warnings).toHaveLength(3);
	});

	it('counts conversions across a large file without one warning per cell', () => {
		const rows = ['id,acct,code,name'];
		for (let i = 0; i < 20_000; i++)
			rows.push(`${String(i + 1).padStart(6, '0')},${i + 1}77777777777777777,${i % 9}E2,n${i}`);
		const csv = load(rows.join('\n'));
		expect(csv.warnings).toEqual([
			'Column A: 20000 values lost their leading zeros when converted to numbers (first: A2).',
			'Column B: 20000 values with more than 15 digits lost digits when converted to numbers (first: B2).',
		]);
		expect(csv.value('A20001')).toBe(20_000);
		const kept = load(rows.join('\n'), { columns: { 0: 'text', 1: 'text', 2: 'text' } });
		expect(kept.warnings).toEqual([]);
		expect(kept.value('A20001')).toBe('020000');
	});
});
