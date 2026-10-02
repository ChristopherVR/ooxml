import { describe, expect, it } from 'vitest';
import { getCell } from '../cells.js';
import { styleAt } from '../styles.js';
import { createWorkbook } from '../workbook.js';
import {
	csvToWorkbook,
	detectDelimiter,
	parseCsv,
	sheetNameFromFileName,
	sheetToCsv,
} from './csv.js';

describe('parseCsv', () => {
	it('reads quoted fields with doubled quotes, delimiters and embedded line breaks', () => {
		const text = 'name,note\r\n"Smith, J","said ""hi"""\r\n"multi\nline",x\r\n';
		expect(parseCsv(text)).toEqual([
			['name', 'note'],
			['Smith, J', 'said "hi"'],
			['multi\nline', 'x'],
		]);
	});

	it('accepts LF, CRLF and CR line endings and keeps empty fields and rows', () => {
		expect(parseCsv('a,b\nc,\r\n\r,d')).toEqual([['a', 'b'], ['c', ''], [''], ['', 'd']]);
		expect(parseCsv('')).toEqual([]);
		expect(parseCsv('only')).toEqual([['only']]);
	});

	it('strips a byte-order mark', () => {
		expect(parseCsv('\ufeffa,b\n1,2\n')).toEqual([
			['a', 'b'],
			['1', '2'],
		]);
	});

	it('honours an explicit delimiter', () => {
		expect(parseCsv('a|b|"c|d"', { delimiter: '|' })).toEqual([['a', 'b', 'c|d']]);
	});
});

describe('detectDelimiter', () => {
	it('picks the delimiter that splits lines consistently', () => {
		expect(detectDelimiter('a,b,c\n1,2,3\n')).toBe(',');
		expect(detectDelimiter('a;b;c\n1;2;3\n')).toBe(';');
		expect(detectDelimiter('a\tb\n1\t2\n')).toBe('\t');
		expect(detectDelimiter('single column\nvalues\n')).toBe(',');
	});

	it('prefers semicolons when commas are decimal separators', () => {
		expect(detectDelimiter('price;qty\n1,5;2\n2,25;3\n')).toBe(';');
		expect(detectDelimiter('"x;y",b\n"p;q",c\n')).toBe(',');
	});
});

describe('csvToWorkbook', () => {
	it('types values through parseCellInput', () => {
		const workbook = csvToWorkbook(
			'Item;Price;When;Ok;Share;F\nTea;1234.5;2024-01-15;TRUE;12%;=B2*2\n',
			{
				sheetName: 'prices',
			},
		);
		const sheet = workbook.sheets[0]!;
		expect(workbook.format).toBe('csv');
		expect(sheet.name).toBe('prices');
		expect(getCell(sheet, 1, 0)?.value).toBe('Tea');
		expect(getCell(sheet, 1, 1)?.value).toBe(1234.5);
		expect(typeof getCell(sheet, 1, 2)?.value).toBe('number');
		expect(styleAt(workbook, getCell(sheet, 1, 2)?.styleId).numFmt).toMatch(/y/);
		expect(getCell(sheet, 1, 3)?.value).toBe(true);
		expect(getCell(sheet, 1, 4)?.value).toBeCloseTo(0.12);
		expect(getCell(sheet, 1, 5)?.formula).toBe('B2*2');
		expect(getCell(sheet, 0, 6)).toBeUndefined();
	});
});

describe('sheetToCsv', () => {
	it('writes formatted values, quotes when needed and pads rows to the used width', () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		const percent = workbook.styles.push({ ...workbook.styles[0]!, numFmt: '0.0%' }) - 1;
		sheet.rows.set(
			0,
			new Map([
				[0, { value: 'a,b' }],
				[1, { value: 'say "x"' }],
				[2, { value: 'line\nbreak' }],
			]),
		);
		sheet.rows.set(
			2,
			new Map([
				[0, { value: 0.125, styleId: percent }],
				[1, { value: true }],
			]),
		);
		expect(sheetToCsv(workbook, 0)).toBe(
			'"a,b","say ""x""","line\nbreak"\r\n,,\r\n12.5%,TRUE,\r\n',
		);
		expect(sheetToCsv(workbook, 0, { delimiter: ';' }).split('\r\n')[2]).toBe('12.5%;TRUE;');
	});

	it('round-trips through csvToWorkbook', () => {
		const source = 'h1,h2\r\n"x, y",2\r\n';
		expect(sheetToCsv(csvToWorkbook(source), 0)).toBe(source);
	});

	it('writes nothing for an empty sheet and rejects a missing one', () => {
		expect(sheetToCsv(createWorkbook(), 0)).toBe('');
		expect(() => sheetToCsv(createWorkbook(), 3)).toThrow(RangeError);
	});
});

describe('sheetNameFromFileName', () => {
	it('derives a valid sheet name', () => {
		expect(sheetNameFromFileName('C:\\data\\Q1 report.csv')).toBe('Q1 report');
		expect(sheetNameFromFileName('a[1]:b.csv')).toBe('a_1__b');
		expect(sheetNameFromFileName(undefined)).toBe('Sheet1');
		expect(sheetNameFromFileName('x'.repeat(40) + '.csv')).toHaveLength(31);
	});
});
