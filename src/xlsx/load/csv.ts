// CSV (RFC 4180) reading and writing for the workbook model: quoted fields with doubled quotes,
// embedded line breaks, CRLF/LF/CR line endings, a leading byte-order mark and delimiter
// detection between comma, semicolon and tab.
import { usedRange } from '../cells.js';
import type { Cell, Workbook } from '../model.js';
import { formatValue, parseCellInput, type ParsedInput } from '../numfmt/index.js';
import { styleAt, internStyle } from '../styles.js';
import { createWorkbook } from '../workbook.js';

export interface CsvOptions {
	/** Field delimiter; detected from the text when omitted. */
	delimiter?: string;
}

const CANDIDATES = [',', ';', '\t'] as const;

/** Counts each candidate delimiter outside quotes, line by line, over the first lines of `text`. */
function delimiterCounts(text: string, maxLines = 20): Map<string, number[]> {
	const counts = new Map<string, number[]>(CANDIDATES.map((d) => [d, [0]]));
	let quoted = false;
	let lines = 1;
	for (let i = 0; i < text.length && lines <= maxLines; i++) {
		const char = text[i];
		if (char === '"') quoted = !quoted;
		else if (!quoted && (char === '\n' || char === '\r')) {
			if (char === '\r' && text[i + 1] === '\n') i++;
			if (i + 1 >= text.length) break;
			lines++;
			for (const list of counts.values()) list.push(0);
		} else if (!quoted) {
			const list = counts.get(char ?? '');
			if (list) list[list.length - 1] = (list[list.length - 1] ?? 0) + 1;
		}
	}
	return counts;
}

/**
 * The delimiter that splits the most lines into the same, non-zero number of fields. On a tie,
 * tab beats semicolon beats comma: a consistent semicolon or tab is the separator, while commas
 * then are usually decimal separators or text.
 */
export function detectDelimiter(text: string): string {
	const preference = ['\t', ';', ','];
	let best = { delimiter: ',', score: 0 };
	for (const [delimiter, list] of delimiterCounts(text)) {
		const tally = new Map<number, number>();
		for (const count of list) if (count > 0) tally.set(count, (tally.get(count) ?? 0) + 1);
		const score = Math.max(0, ...tally.values());
		const better =
			score > best.score ||
			(score === best.score &&
				score > 0 &&
				preference.indexOf(delimiter) < preference.indexOf(best.delimiter));
		if (better) best = { delimiter, score };
	}
	return best.delimiter;
}

/** Removes a UTF-8 byte-order mark decoded as U+FEFF. */
const stripBom = (text: string): string => (text.charCodeAt(0) === 0xfeff ? text.slice(1) : text);

/** Parses CSV text into rows of fields. A trailing line break does not add an empty row. */
export function parseCsv(input: string, options: CsvOptions = {}): string[][] {
	const text = stripBom(input);
	const delimiter = options.delimiter ?? detectDelimiter(text);
	const rows: string[][] = [];
	let row: string[] = [];
	let field = '';
	let quoted = false;
	let fieldStarted = false;
	for (let i = 0; i < text.length; i++) {
		const char = text[i] ?? '';
		if (quoted) {
			if (char === '"') {
				if (text[i + 1] === '"') {
					field += '"';
					i++;
				} else quoted = false;
			} else field += char;
			continue;
		}
		if (char === '"' && !fieldStarted) {
			quoted = true;
			fieldStarted = true;
		} else if (text.startsWith(delimiter, i)) {
			row.push(field);
			field = '';
			fieldStarted = false;
			i += delimiter.length - 1;
		} else if (char === '\n' || char === '\r') {
			if (char === '\r' && text[i + 1] === '\n') i++;
			row.push(field);
			rows.push(row);
			row = [];
			field = '';
			fieldStarted = false;
		} else {
			field += char;
			fieldStarted = true;
		}
	}
	if (fieldStarted || field !== '' || row.length > 0) {
		row.push(field);
		rows.push(row);
	}
	return rows;
}

/** Characters Excel does not allow in a sheet name, replaced when a file name becomes one. */
const INVALID_SHEET_CHARS = /[\\/?*[\]:]/g;

/** A valid sheet name derived from a file name (`report.csv` -> `report`). */
export function sheetNameFromFileName(fileName: string | undefined): string {
	const base = (fileName ?? '').split(/[\\/]/).pop() ?? '';
	const stem = base.replace(/\.[^.]*$/, '').replace(INVALID_SHEET_CHARS, '_');
	const trimmed = stem
		.replace(/^'+|'+$/g, '')
		.slice(0, 31)
		.trim();
	return trimmed && trimmed.toLowerCase() !== 'history' ? trimmed : 'Sheet1';
}

export interface CsvWorkbookOptions extends CsvOptions {
	sheetName?: string;
	/**
	 * Treat fields starting with `=` as formulas, as Excel does when it opens a CSV. Off by
	 * default: a CSV from an untrusted source must not inject formulas (`=HYPERLINK(...)`,
	 * `=WEBSERVICE(...)`) into the workbook, so such fields are imported as text.
	 */
	formulas?: boolean;
}

/**
 * A one-sheet workbook from CSV text. Each field goes through `parseCellInput`, so numbers,
 * dates, percentages, booleans and errors become typed cells as Excel opens them; `=` fields
 * become formulas only with `formulas: true` and are text otherwise.
 */
export function csvToWorkbook(text: string, options: CsvWorkbookOptions = {}): Workbook {
	const workbook = createWorkbook({ sheets: [options.sheetName ?? 'Sheet1'] });
	workbook.format = 'csv';
	const sheet = workbook.sheets[0];
	if (!sheet) return workbook;
	const base = styleAt(workbook, 0);
	const formats = new Map<string, number>();
	const rows = parseCsv(
		text,
		options.delimiter === undefined ? {} : { delimiter: options.delimiter },
	);
	rows.forEach((fields, row) => {
		fields.forEach((field, col) => {
			if (field === '') return;
			const parsed: ParsedInput =
				!options.formulas && field.startsWith('=') ? { value: field } : parseCellInput(field);
			const cell: Cell = { value: parsed.value };
			if (parsed.formula !== undefined) cell.formula = parsed.formula;
			if (parsed.numFmt && parsed.numFmt !== 'General') {
				let styleId = formats.get(parsed.numFmt);
				if (styleId === undefined) {
					styleId = internStyle(workbook, { ...base, numFmt: parsed.numFmt });
					formats.set(parsed.numFmt, styleId);
				}
				if (styleId) cell.styleId = styleId;
			}
			let cells = sheet.rows.get(row);
			if (!cells) {
				cells = new Map();
				sheet.rows.set(row, cells);
			}
			cells.set(col, cell);
		});
	});
	return workbook;
}

/** Leading characters a spreadsheet may read as the start of a formula (OWASP CSV injection). */
const FORMULA_TRIGGER = /^[=+\-@\t\r]/;

/**
 * A text value that would run as a formula when the CSV is opened in a spreadsheet, prefixed
 * with a single quote so it stays text (OWASP CSV injection guidance). Numbers such as `-5` are
 * not text cells and are never prefixed.
 */
export function neutraliseCsvFormula(text: string): string {
	return FORMULA_TRIGGER.test(text) ? `'${text}` : text;
}

export interface CsvWriteOptions extends CsvOptions {
	/**
	 * Prefix text that starts with `=`, `+`, `-`, `@`, tab or carriage return with `'` so a
	 * spreadsheet opening the file does not evaluate it. On by default.
	 */
	escapeFormulas?: boolean;
}

function quoteField(text: string, delimiter: string): string {
	return text.includes(delimiter) || /["\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

/**
 * One sheet as CSV: each cell's formatted display text (as Excel's CSV export writes it), rows
 * from A1 to the end of the used range padded to the same width, CRLF line endings. Text that a
 * spreadsheet would evaluate as a formula is neutralised (see {@link neutraliseCsvFormula}).
 */
export function sheetToCsv(
	workbook: Workbook,
	sheetIndex: number,
	options: CsvWriteOptions = {},
): string {
	const escape = options.escapeFormulas ?? true;
	const sheet = workbook.sheets[sheetIndex];
	if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
	const delimiter = options.delimiter ?? ',';
	const range = usedRange(sheet);
	if (!range) return '';
	const lines: string[] = [];
	for (let row = 0; row <= range.end.row; row++) {
		const cells = sheet.rows.get(row);
		const fields: string[] = [];
		for (let col = 0; col <= range.end.col; col++) {
			const cell = cells?.get(col);
			let text = cell
				? formatValue(cell.value, styleAt(workbook, cell.styleId).numFmt, {
						date1904: workbook.date1904,
					}).text
				: '';
			if (escape && typeof cell?.value === 'string') text = neutraliseCsvFormula(text);
			fields.push(quoteField(text, delimiter));
		}
		lines.push(fields.join(delimiter));
	}
	return `${lines.join('\r\n')}\r\n`;
}
