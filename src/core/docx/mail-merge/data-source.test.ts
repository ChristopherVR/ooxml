import { describe, expect, it } from 'vitest';
import {
	createWorkbook,
	dateToSerial,
	internStyle,
	putCell,
	saveXlsx,
	styleAt,
} from '../../xlsx/index';
import {
	mailMergeDelimitedChunks,
	mailMergeDelimitedText,
	mailMergeRecords,
	mailMergeXlsx,
	selectMailMergeRecipients,
} from './index';

describe('mail merge recipient data', () => {
	it('reads quoted CSV fields with line breaks, BOM, leading zeroes and literal formulas', async () => {
		const source = mailMergeDelimitedText(
			'\uFEFFName,Postal,Notes\r\n"Doe, Jane",00120,"line 1\nline ""2"""\r\nLee,,=1+1',
		);
		expect(await source.read()).toEqual({
			fields: ['Name', 'Postal', 'Notes'],
			records: [
				{ Name: 'Doe, Jane', Postal: '00120', Notes: 'line 1\nline "2"' },
				{ Name: 'Lee', Postal: '', Notes: '=1+1' },
			],
		});
	});

	it('reads TSV and filters and sorts a host record list without mutating it', async () => {
		const input = await mailMergeDelimitedText(
			'Name\tCity\nLee\tParis\nAda\tRome\nAri\tParis\nBo\tParis',
		).read();
		const provider = mailMergeRecords(input);
		const data = await provider.read();
		const selected = selectMailMergeRecipients(data, {
			filter: (record, index) => record.City === 'Paris' && index !== 3,
			sort: (left, right) => String(left.Name).localeCompare(String(right.Name)),
		});
		expect(selected.fields).toEqual(['Name', 'City']);
		expect(selected.records.map(({ Name }) => Name)).toEqual(['Ari', 'Lee']);
		expect(data.records.map(({ Name }) => Name)).toEqual(['Lee', 'Ada', 'Ari', 'Bo']);
		expect(() => mailMergeDelimitedText('Name, Name\nA,B').read()).toThrow(
			'Duplicate mail merge field',
		);
	});

	it('reads a named XLSX sheet, typed values, 1904 dates and cached formulas', async () => {
		const book = createWorkbook({ sheets: ['Other', 'Recipients'] });
		const sheet = book.sheets[1]!;
		book.date1904 = true;
		const styleId = internStyle(book, { ...styleAt(book, 0), numFmt: 'yyyy-mm-dd' });
		['Name', 'Birthday', '', 'Total', 'Active'].forEach((value, col) =>
			putCell(sheet, 0, col, { value }),
		);
		putCell(sheet, 1, 0, { value: 'Ada' });
		putCell(sheet, 1, 1, { value: dateToSerial(new Date('2000-12-10T00:00Z'), true), styleId });
		putCell(sheet, 1, 3, { value: 42, formula: '40+2' });
		putCell(sheet, 1, 4, { value: true });
		putCell(sheet, 4, 0, { value: 'Bea' });
		const bytes = await saveXlsx(book);
		expect(await mailMergeXlsx(bytes, { sheet: 'recipients' }).read()).toEqual({
			fields: ['Name', 'Birthday', 'Total', 'Active'],
			records: [
				{ Name: 'Ada', Birthday: '2000-12-10', Total: 42, Active: true },
				{ Name: 'Bea', Birthday: null, Total: null, Active: null },
			],
		});
		await expect(mailMergeXlsx(bytes, { sheet: 'Missing' }).read()).rejects.toThrow('Missing');
	});

	it('reads delimited text chunk by chunk with the same records and bounded limits', async () => {
		const text = 'Name,Postal\r\n"Doe, Jane",00120\r\nLee,"a\nb"\r\n';
		const whole = await mailMergeDelimitedText(text).read();
		const chunks = await mailMergeDelimitedChunks(() => [...text]).read();
		expect(chunks).toEqual(whole);
		async function* slow() {
			yield 'Name\nAda\n';
			yield 'Bea\nCal\n';
		}
		const limits = { maxFields: 4, maxRecords: 2 };
		await expect(mailMergeDelimitedChunks(slow, { dataLimits: limits }).read()).rejects.toThrow(
			RangeError,
		);
		const long = { maxFields: 4, maxRecords: 10, maxValueLength: 3 };
		await expect(mailMergeDelimitedChunks(slow, { dataLimits: long }).read()).resolves.toEqual({
			fields: ['Name'],
			records: [{ Name: 'Ada' }, { Name: 'Bea' }, { Name: 'Cal' }],
		});
		const tight = { maxFields: 4, maxRecords: 10, maxCharacters: 8 };
		await expect(mailMergeDelimitedChunks(slow, { dataLimits: tight }).read()).rejects.toThrow(
			'too large',
		);
		const cells = { maxFields: 4, maxRecords: 10, maxCells: 2 };
		await expect(mailMergeDelimitedChunks(slow, { dataLimits: cells }).read()).rejects.toThrow(
			'values',
		);
	});
});
