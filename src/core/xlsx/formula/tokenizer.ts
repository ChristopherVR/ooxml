import { FormulaError, type RefSpec, type SheetPrefix } from './ast.js';
import {
	NAME_RE,
	NAME_START,
	NUMBER_RE,
	OPERATORS,
	matchError,
	readBrackets,
	readPrefix,
	readReference,
	numberLiteral,
} from './lexemes.js';

export { numberLiteral, readReference } from './lexemes.js';

export type TokenKind =
	| 'number'
	| 'string'
	| 'bool'
	| 'error'
	| 'ref'
	| 'name'
	| 'func'
	| 'structured'
	| 'op'
	| 'open'
	| 'close'
	| 'comma'
	| 'semicolon'
	| 'lbrace'
	| 'rbrace'
	| 'ws';

/** A lexical token. `text` is the exact source slice, so joining every token's text round-trips. */
export interface Token {
	kind: TokenKind;
	text: string;
	start: number;
	/** Number value, unescaped string, boolean, error code, or name / function / table name. */
	value?: number | string | boolean;
	prefix?: SheetPrefix;
	/** Parsed reference (`ref` tokens); absent for `#REF!`. */
	ref?: RefSpec;
	/** `A1#` spill reference. */
	spill?: boolean;
}

/** Splits a formula (with or without its leading `=`) into tokens, whitespace included. */
export function tokenize(formula: string): Token[] {
	const source = formula;
	const tokens: Token[] = [];
	let i = source.startsWith('=') ? 1 : 0;
	if (i === 1) tokens.push({ kind: 'ws', text: '=', start: 0 });
	const push = (token: Token): void => {
		tokens.push(token);
		i = token.start + token.text.length;
	};
	while (i < source.length) {
		const ch = source[i] ?? '';
		const start = i;
		if (/[ \t\r\n]/.test(ch)) {
			const m = /^[ \t\r\n]+/.exec(source.slice(i));
			push({ kind: 'ws', text: m?.[0] ?? ch, start });
			continue;
		}
		if (ch === '"') {
			let j = i + 1;
			let value = '';
			for (;;) {
				if (j >= source.length) throw new FormulaError('Unterminated string', start);
				if (source[j] === '"') {
					if (source[j + 1] === '"') {
						value += '"';
						j += 2;
						continue;
					}
					break;
				}
				value += source[j];
				j++;
			}
			push({ kind: 'string', text: source.slice(start, j + 1), start, value });
			continue;
		}
		if (ch === '#') {
			const code = matchError(source, i);
			if (!code) throw new FormulaError(`Unexpected '#'`, start);
			push({
				kind: code === '#REF!' ? 'ref' : 'error',
				text: source.slice(i, i + code.length),
				start,
				value: code,
			});
			continue;
		}
		const single: Partial<Record<string, TokenKind>> = {
			'(': 'open',
			')': 'close',
			',': 'comma',
			';': 'semicolon',
			'{': 'lbrace',
			'}': 'rbrace',
		};
		const kind = single[ch];
		if (kind) {
			push({ kind, text: ch, start });
			continue;
		}
		if (ch === '[' && !/^\[\d+\]/.test(source.slice(i))) {
			const length = readBrackets(source, i);
			push({ kind: 'structured', text: source.slice(i, i + length), start, value: '' });
			continue;
		}
		const prefixed = ch === "'" || ch === '[' || NAME_START.test(ch) || /\d/.test(ch);
		if (prefixed) {
			const prefix = readPrefix(source, i);
			if (prefix) {
				push(readAfterPrefix(source, i + prefix.length, start, prefix.prefix));
				continue;
			}
			if (ch === "'" || ch === '[') throw new FormulaError('Malformed sheet reference', start);
		}
		if (/[\d.]/.test(ch) || ch === '$' || NAME_START.test(ch)) {
			const ref = readReference(source, i);
			if (ref) {
				push(refToken(source, start, i + ref.length, ref.ref, undefined));
				continue;
			}
			if (/[\d.]/.test(ch)) {
				const m = NUMBER_RE.exec(source.slice(i));
				if (!m) throw new FormulaError(`Unexpected '${ch}'`, start);
				push({ kind: 'number', text: m[0], start, value: numberLiteral(m[0]) });
				continue;
			}
			if (ch !== '$') {
				push(readName(source, i, start, undefined));
				continue;
			}
		}
		const op = OPERATORS.find((o) => source.startsWith(o, i));
		if (op) {
			push({ kind: 'op', text: op, start });
			continue;
		}
		throw new FormulaError(`Unexpected character '${ch}'`, start);
	}
	return tokens;
}

function refToken(
	source: string,
	start: number,
	end: number,
	ref: RefSpec,
	prefix: SheetPrefix | undefined,
): Token {
	let stop = end;
	let spill = false;
	if (ref.kind === 'cell' && source[end] === '#' && !matchError(source, end)) {
		spill = true;
		stop = end + 1;
	}
	const token: Token = { kind: 'ref', text: source.slice(start, stop), start, ref };
	if (prefix) token.prefix = prefix;
	if (spill) token.spill = true;
	return token;
}

function readAfterPrefix(source: string, at: number, start: number, prefix: SheetPrefix): Token {
	if (source.slice(at, at + 5).toUpperCase() === '#REF!') {
		return { kind: 'ref', text: source.slice(start, at + 5), start, prefix };
	}
	const ref = readReference(source, at);
	if (ref) return refToken(source, start, at + ref.length, ref.ref, prefix);
	if (NAME_START.test(source[at] ?? '')) return readName(source, at, start, prefix);
	throw new FormulaError('Invalid reference after sheet name', at);
}

function readName(
	source: string,
	at: number,
	start: number,
	prefix: SheetPrefix | undefined,
): Token {
	const m = NAME_RE.exec(source.slice(at));
	if (!m) throw new FormulaError('Invalid name', at);
	const name = m[0];
	let end = at + name.length;
	if (source[end] === '[' && !prefix) {
		end += readBrackets(source, end);
		return { kind: 'structured', text: source.slice(start, end), start, value: name };
	}
	if (source[end] === '(') {
		return { kind: 'func', text: source.slice(start, end), start, value: name };
	}
	const upper = name.toUpperCase();
	if (!prefix && (upper === 'TRUE' || upper === 'FALSE')) {
		return { kind: 'bool', text: name, start, value: upper === 'TRUE' };
	}
	const token: Token = { kind: 'name', text: source.slice(start, end), start, value: name };
	if (prefix) token.prefix = prefix;
	return token;
}

/** Reassembles tokens into formula text. */
export const joinTokens = (tokens: readonly Token[]): string => tokens.map((t) => t.text).join('');
