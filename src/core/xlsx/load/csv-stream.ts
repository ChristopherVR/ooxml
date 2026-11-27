// Incremental CSV (RFC 4180) reading: the same rules as `parseCsv` (quoted fields with doubled
// quotes, embedded line breaks, CRLF/LF/CR, a leading byte-order mark, delimiter detection), fed
// one chunk of text at a time so a large file never has to be held as one string.
import { detectDelimiter } from './csv';

export interface CsvRowReaderOptions {
	/** A one-character field delimiter; detected from the first lines when omitted. */
	delimiter?: string;
	maxRows?: number;
	maxColumns?: number;
	/** The longest field, in UTF-16 code units. */
	maxFieldLength?: number;
}

export interface CsvRowReader {
	/** Reads a chunk and returns the rows it completed. */
	push(chunk: string): string[][];
	/** Ends the input and returns the rows still pending. */
	end(): string[][];
}

/** Text buffered before the delimiter is detected: 20 lines or this many code units. */
const DETECT_LENGTH = 64 * 1024;

export function csvRowReader(options: CsvRowReaderOptions = {}): CsvRowReader {
	if (options.delimiter !== undefined && options.delimiter.length !== 1)
		throw new RangeError('The delimiter must be one character.');
	let delimiter = options.delimiter;
	let pending = '';
	let first = true;
	let rowCount = 0;
	let row: string[] = [];
	let field = '';
	let quoted = false;
	let fieldStarted = false;
	/** A quote inside a quoted field whose meaning depends on the next character. */
	let quote = false;
	/** A carriage return whose line feed may start the next chunk. */
	let carriage = false;
	let output: string[][] = [];
	const maxField = options.maxFieldLength ?? Infinity;

	const append = (text: string) => {
		field += text;
		if (field.length > maxField) throw new RangeError('A delimited text field is too long.');
	};
	const pushField = () => {
		if (options.maxColumns !== undefined && row.length >= options.maxColumns)
			throw new RangeError('Too many delimited text columns.');
		row.push(field);
		field = '';
		fieldStarted = false;
	};
	const pushRow = () => {
		if (options.maxRows !== undefined && rowCount >= options.maxRows)
			throw new RangeError('Too many delimited text rows.');
		pushField();
		output.push(row);
		rowCount++;
		row = [];
	};
	const consume = (text: string) => {
		for (let i = 0; i < text.length; i++) {
			const char = text[i]!;
			if (carriage) {
				carriage = false;
				if (char === '\n') continue;
			}
			if (quote) {
				quote = false;
				if (char === '"') {
					append('"');
					continue;
				}
				quoted = false;
			}
			if (quoted) {
				const close = text.indexOf('"', i);
				if (close < 0) {
					append(text.slice(i));
					return;
				}
				append(text.slice(i, close));
				i = close;
				quote = true;
				continue;
			}
			if (char === '"' && !fieldStarted) {
				quoted = true;
				fieldStarted = true;
			} else if (char === delimiter) pushField();
			else if (char === '\n' || char === '\r') {
				carriage = char === '\r';
				pushRow();
			} else {
				append(char);
				fieldStarted = true;
			}
		}
	};
	const take = () => {
		const rows = output;
		output = [];
		return rows;
	};
	const lineBreaks = (text: string) => {
		let count = 0;
		for (let i = 0; i < text.length && count < 21; i++) if (text[i] === '\n') count++;
		return count;
	};

	return {
		push(chunk) {
			if (first && chunk) {
				first = false;
				if (chunk.charCodeAt(0) === 0xfeff) chunk = chunk.slice(1);
			}
			if (delimiter === undefined) {
				pending += chunk;
				if (pending.length < DETECT_LENGTH && lineBreaks(pending) <= 20) return [];
				delimiter = detectDelimiter(pending);
				chunk = pending;
				pending = '';
			}
			consume(chunk);
			return take();
		},
		end() {
			if (delimiter === undefined) {
				delimiter = detectDelimiter(pending);
				consume(pending);
				pending = '';
			}
			if (quote) quoted = false;
			quote = false;
			if (fieldStarted || field !== '' || row.length > 0) pushRow();
			return take();
		},
	};
}
