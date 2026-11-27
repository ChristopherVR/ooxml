import { describe, expect, it } from 'vitest';
import { parseCsv } from './csv';
import { csvRowReader, type CsvRowReaderOptions } from './csv-stream';

function readChunks(chunks: readonly string[], options?: CsvRowReaderOptions): string[][] {
	const reader = csvRowReader(options);
	return [...chunks.flatMap((chunk) => reader.push(chunk)), ...reader.end()];
}

/** Every way of cutting `text` into two chunks, plus one chunk per character. */
function splits(text: string): string[][] {
	const cuts = Array.from({ length: text.length + 1 }, (_, at) => [
		text.slice(0, at),
		text.slice(at),
	]);
	return [...cuts, [...text]];
}

const SAMPLES = [
	'Name,City\r\nAda,London\r\nBea,"Paris, France"\r\n',
	'﻿a;b;c\n1;2;3\n',
	'Name\tNote\n"Ann"\t"She said ""hi""\nthen left"\nBob\t\r\nCal\t"x"y\r',
	'a,b\n"unterminated,field\nmore',
	'only header',
	'a,b\n\n,\n',
];

describe('incremental CSV reader', () => {
	it('reads the same rows as parseCsv however the text is cut', () => {
		for (const sample of SAMPLES)
			for (const chunks of splits(sample))
				expect(readChunks(chunks), sample).toEqual(parseCsv(sample));
	});

	it('detects the delimiter from the first lines before reading rows', () => {
		const lines = Array.from({ length: 30 }, (_, i) => `${i};x;y`).join('\n');
		const reader = csvRowReader();
		expect(reader.push(lines.slice(0, 20))).toEqual([]);
		const rows = [...reader.push(lines.slice(20)), ...reader.end()];
		expect(rows).toHaveLength(30);
		expect(rows[0]).toEqual(['0', 'x', 'y']);
	});

	it('enforces row, column and field length limits while reading', () => {
		expect(() => readChunks(['a\n1\n2\n'], { delimiter: ',', maxRows: 2 })).toThrow(RangeError);
		expect(() => readChunks(['a,b,c\n'], { delimiter: ',', maxColumns: 2 })).toThrow(RangeError);
		expect(() => readChunks(['"abc', 'def"\n'], { delimiter: ',', maxFieldLength: 4 })).toThrow(
			RangeError,
		);
		expect(() => csvRowReader({ delimiter: '::' })).toThrow(RangeError);
	});
});
