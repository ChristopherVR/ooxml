// Lexical building blocks of the tokenizer: names, references, sheet prefixes, literals.
import { columnIndex, MAX_COL, MAX_ROW } from '../address';
import { ERROR_CODES, type ErrorCode } from '../model';
import { FormulaError, type RefCorner, type RefSpec, type SheetPrefix } from './ast';

export const ERRORS_BY_LENGTH = [...ERROR_CODES].sort((a, b) => b.length - a.length);
export const NAME_START = /[A-Za-z_\\\u00A1-\uFFFF]/;
export const NAME_CHAR = /[A-Za-z0-9_.?\\\u00A1-\uFFFF]/;
export const NAME_RE = /^[A-Za-z_\\\u00A1-\uFFFF][A-Za-z0-9_.?\\\u00A1-\uFFFF]*/;
export const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;
export const COLS_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/;
export const ROWS_RE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})/;
export const NUMBER_RE = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/;

/** Sticky forms of the anchored patterns, matched at a position without slicing the source. */
const sticky = (re: RegExp): RegExp => new RegExp(re.source.replace(/^\^/, ''), 'y');
const CELL_AT = sticky(CELL_RE);
const COLS_AT = sticky(COLS_RE);
const ROWS_AT = sticky(ROWS_RE);
export const NAME_AT = sticky(NAME_RE);
export const NUMBER_AT = sticky(NUMBER_RE);

/** Matches a sticky pattern at `at`. */
export function matchAt(re: RegExp, source: string, at: number): RegExpExecArray | null {
	re.lastIndex = at;
	return re.exec(source);
}

const isDigitAt = (source: string, at: number): boolean => isDigitCode(source.charCodeAt(at));

/**
 * Whether `[$]run:` can start at `at`, with a run of 1 to `max` characters accepted by `accept`:
 * the shape `COLS_RE` and `ROWS_RE` need before they can match, checked without a regex.
 */
function rangeStartAt(
	source: string,
	at: number,
	accept: (code: number) => boolean,
	max: number,
): boolean {
	let i = source[at] === '$' ? at + 1 : at;
	const from = i;
	while (i - from <= max && accept(source.charCodeAt(i))) i++;
	return i > from && i - from <= max && source[i] === ':';
}

export const OPERATORS = [
	'<=',
	'>=',
	'<>',
	'<',
	'>',
	'=',
	'+',
	'-',
	'*',
	'/',
	'^',
	'&',
	'%',
	':',
	'@',
];

/** `[A-Za-z]` by character code. */
const isLetterCode = (code: number): boolean =>
	(code >= 65 && code <= 90) || (code >= 97 && code <= 122);
const isDigitCode = (code: number): boolean => code >= 48 && code <= 57;

/** {@link NAME_START} by character code (its last range covers every code unit from 0xA1). */
export const isNameStartCode = (code: number): boolean =>
	isLetterCode(code) || code === 95 || code === 92 || code >= 0xa1;

/** {@link NAME_CHAR} by character code. */
const isNameCharCode = (code: number): boolean =>
	isNameStartCode(code) || isDigitCode(code) || code === 46 || code === 63;

/** A character of an unquoted sheet name in a prefix (the class of `PREFIX_AT`). */
const isSheetCharCode = (code: number): boolean =>
	isLetterCode(code) || isDigitCode(code) || code === 95 || code === 46 || code >= 0xa1;

export const isNameChar = (ch: string | undefined): boolean =>
	ch !== undefined && ch !== '' && isNameCharCode(ch.charCodeAt(0));

export function matchError(source: string, at: number): ErrorCode | undefined {
	const upper = source.slice(at, at + 14).toUpperCase();
	return ERRORS_BY_LENGTH.find((code) => upper.startsWith(code));
}

export function corner(
	colAbs: string,
	col: string,
	rowAbs: string,
	row: string,
): RefCorner | undefined {
	const c = columnIndex(col);
	const r = Number(row) - 1;
	if (c < 0 || c > MAX_COL || r < 0 || r > MAX_ROW) return undefined;
	return { row: r, col: c, rowAbs: rowAbs === '$', colAbs: colAbs === '$' };
}

/** Whether the reference text ending at `end` is not followed by more identifier characters. */
export const boundary = (source: string, end: number): boolean => {
	const next = source[end];
	return !isNameChar(next) && next !== '(' && next !== '!' && next !== '[';
};

/** Tries to read an A1 reference at `at`; returns its spec and length. */
export function readReference(
	source: string,
	at: number,
): { ref: RefSpec; length: number } | undefined {
	const cell = matchAt(CELL_AT, source, at);
	if (cell) {
		const start = corner(cell[1] ?? '', cell[2] ?? '', cell[3] ?? '', cell[4] ?? '');
		let length = cell[0].length;
		if (start) {
			if (source[at + length] === ':') {
				const second = matchAt(CELL_AT, source, at + length + 1);
				if (second) {
					const end = corner(second[1] ?? '', second[2] ?? '', second[3] ?? '', second[4] ?? '');
					const total = length + 1 + second[0].length;
					if (end && boundary(source, at + total)) {
						return { ref: { kind: 'area', start, end }, length: total };
					}
				}
			}
			if (boundary(source, at + length))
				return { ref: { kind: 'cell', start, end: start }, length };
		}
		length = 0;
	}
	const cols = rangeStartAt(source, at, isLetterCode, 3) ? matchAt(COLS_AT, source, at) : null;
	if (cols && boundary(source, at + cols[0].length) && !isDigitAt(source, at + cols[0].length)) {
		const c1 = columnIndex(cols[2] ?? '');
		const c2 = columnIndex(cols[4] ?? '');
		if (c1 >= 0 && c1 <= MAX_COL && c2 >= 0 && c2 <= MAX_COL) {
			return {
				ref: {
					kind: 'cols',
					start: { row: 0, col: c1, rowAbs: false, colAbs: cols[1] === '$' },
					end: { row: MAX_ROW, col: c2, rowAbs: false, colAbs: cols[3] === '$' },
				},
				length: cols[0].length,
			};
		}
	}
	const rows = rangeStartAt(source, at, isDigitCode, 7) ? matchAt(ROWS_AT, source, at) : null;
	if (rows && boundary(source, at + rows[0].length) && source[at + rows[0].length] !== '.') {
		const r1 = Number(rows[2]) - 1;
		const r2 = Number(rows[4]) - 1;
		if (r1 >= 0 && r1 <= MAX_ROW && r2 >= 0 && r2 <= MAX_ROW) {
			return {
				ref: {
					kind: 'rows',
					start: { row: r1, col: 0, rowAbs: rows[1] === '$', colAbs: false },
					end: { row: r2, col: MAX_COL, rowAbs: rows[3] === '$', colAbs: false },
				},
				length: rows[0].length,
			};
		}
	}
	return undefined;
}

/** Reads a `'quoted name'` starting at `at`; returns the unescaped name and the length. */
export function readQuoted(source: string, at: number): { name: string; length: number } {
	let i = at + 1;
	let name = '';
	while (i < source.length) {
		const ch = source[i];
		if (ch === "'") {
			if (source[i + 1] === "'") {
				name += "'";
				i += 2;
				continue;
			}
			return { name, length: i + 1 - at };
		}
		name += ch;
		i++;
	}
	throw new FormulaError('Unterminated quoted sheet name', at);
}

/** Reads balanced `[...]` brackets starting at `at` (`'` escapes the next character). */
export function readBrackets(source: string, at: number): number {
	let depth = 0;
	let i = at;
	while (i < source.length) {
		const ch = source[i];
		if (ch === "'") {
			i += 2;
			continue;
		}
		if (ch === '[') depth++;
		else if (ch === ']') {
			depth--;
			if (depth === 0) return i + 1 - at;
		}
		i++;
	}
	throw new FormulaError('Unterminated structured reference', at);
}

export function splitSheets(text: string): { sheet: string; sheet2?: string; book?: string } {
	let rest = text;
	let book: string | undefined;
	const bookMatch = /^\[([^\]]*)\]/.exec(rest);
	if (bookMatch) {
		book = bookMatch[1] ?? '';
		rest = rest.slice(bookMatch[0].length);
	}
	const colon = rest.indexOf(':');
	const out: { sheet: string; sheet2?: string; book?: string } = {
		sheet: colon >= 0 ? rest.slice(0, colon) : rest,
	};
	if (colon >= 0) out.sheet2 = rest.slice(colon + 1);
	if (book !== undefined) out.book = book;
	return out;
}

const PREFIX_AT = /(\[\d+\])?([A-Za-z0-9_.\u00A1-\uFFFF]+)(?::([A-Za-z0-9_.\u00A1-\uFFFF]+))?!/y;

/** Reads an optional sheet prefix at `at`; returns it and its length. */
export function readPrefix(
	source: string,
	at: number,
): { prefix: SheetPrefix; length: number } | undefined {
	const ch = source[at];
	if (ch === "'") {
		const quoted = readQuoted(source, at);
		if (source[at + quoted.length] !== '!') return undefined;
		const length = quoted.length + 1;
		return { prefix: { text: source.slice(at, at + length), ...splitSheets(quoted.name) }, length };
	}
	// The pattern can only match when the run of sheet-name characters (after an optional `[n]`
	// book index) ends at `!` or `:`; checking that first skips the regex for most tokens.
	let i = at;
	if (ch === '[') {
		i++;
		while (isDigitCode(source.charCodeAt(i))) i++;
		if (i === at + 1 || source[i] !== ']') return undefined;
		i++;
	}
	while (i < source.length && isSheetCharCode(source.charCodeAt(i))) i++;
	if (source[i] !== '!' && source[i] !== ':') return undefined;
	const m = matchAt(PREFIX_AT, source, at);
	if (!m) return undefined;
	if (!m[1] && /^\d/.test(m[2] ?? '') && !m[3]) return undefined;
	const raw = m[0].slice(0, -1);
	return { prefix: { text: m[0], ...splitSheets(raw) }, length: m[0].length };
}

/** A number literal's value; Excel keeps only 15 significant digits of what was typed. */
export function numberLiteral(text: string): number {
	// At most 15 characters cannot hold more than 15 significant digits.
	if (text.length <= 15) return Number(text);
	const [mantissa = '', exponent] = text.toLowerCase().split('e');
	const digits = mantissa.replace('.', '');
	const firstSignificant = digits.search(/[1-9]/);
	if (firstSignificant < 0 || digits.length - firstSignificant <= 15) return Number(text);
	const point = mantissa.includes('.') ? mantissa.indexOf('.') : mantissa.length;
	const kept = digits.slice(0, firstSignificant + 15);
	const truncated =
		kept.length >= point
			? `${kept.slice(0, point)}.${kept.slice(point)}`
			: kept + '0'.repeat(point - kept.length);
	return Number(`${truncated}${exponent !== undefined ? `e${exponent}` : ''}`);
}
