// Lexical building blocks of the tokenizer: names, references, sheet prefixes, literals.
import { columnIndex, MAX_COL, MAX_ROW } from '../address.js';
import { ERROR_CODES, type ErrorCode } from '../model.js';
import { FormulaError, type RefCorner, type RefSpec, type SheetPrefix } from './ast.js';

export const ERRORS_BY_LENGTH = [...ERROR_CODES].sort((a, b) => b.length - a.length);
export const NAME_START = /[A-Za-z_\\\u00A1-\uFFFF]/;
export const NAME_CHAR = /[A-Za-z0-9_.?\\\u00A1-\uFFFF]/;
export const NAME_RE = /^[A-Za-z_\\\u00A1-\uFFFF][A-Za-z0-9_.?\\\u00A1-\uFFFF]*/;
export const CELL_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/;
export const COLS_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/;
export const ROWS_RE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})/;
export const NUMBER_RE = /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/;
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

export const isNameChar = (ch: string | undefined): boolean =>
	ch !== undefined && NAME_CHAR.test(ch);

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
	const rest = source.slice(at);
	const cell = CELL_RE.exec(rest);
	if (cell) {
		const start = corner(cell[1] ?? '', cell[2] ?? '', cell[3] ?? '', cell[4] ?? '');
		let length = cell[0].length;
		if (start) {
			if (rest[length] === ':') {
				const second = CELL_RE.exec(rest.slice(length + 1));
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
	const cols = COLS_RE.exec(rest);
	if (cols && boundary(source, at + cols[0].length) && !/^\d/.test(rest.slice(cols[0].length))) {
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
	const rows = ROWS_RE.exec(rest);
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
	const m = /^(\[\d+\])?([A-Za-z0-9_.\u00A1-\uFFFF]+)(?::([A-Za-z0-9_.\u00A1-\uFFFF]+))?!/.exec(
		source.slice(at),
	);
	if (!m) return undefined;
	if (!m[1] && /^\d/.test(m[2] ?? '') && !m[3]) return undefined;
	const raw = m[0].slice(0, -1);
	return { prefix: { text: m[0], ...splitSheets(raw) }, length: m[0].length };
}

/** A number literal's value; Excel keeps only 15 significant digits of what was typed. */
export function numberLiteral(text: string): number {
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
