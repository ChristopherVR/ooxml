import type { ErrorCode } from '../model.js';
import {
	type ArrayConstant,
	type BinaryOperator,
	FormulaError,
	type FormulaAst,
	normalizeFunctionName,
} from './ast.js';
import { parseStructured } from './structured.js';
import { type Token, tokenize } from './tokenizer.js';

interface Cursor {
	tokens: Token[];
	/** For each significant token: whether whitespace preceded it. */
	wsBefore: boolean[];
	pos: number;
}

const BINARY: Partial<Record<string, { prec: number; op: BinaryOperator }>> = {
	'=': { prec: 1, op: '=' },
	'<>': { prec: 1, op: '<>' },
	'<': { prec: 1, op: '<' },
	'>': { prec: 1, op: '>' },
	'<=': { prec: 1, op: '<=' },
	'>=': { prec: 1, op: '>=' },
	'&': { prec: 2, op: '&' },
	'+': { prec: 3, op: '+' },
	'-': { prec: 3, op: '-' },
	'*': { prec: 4, op: '*' },
	'/': { prec: 4, op: '/' },
	'^': { prec: 5, op: '^' },
	':': { prec: 10, op: ':' },
};
const PREC_UNARY = 6;
const PREC_PERCENT = 7;
const PREC_UNION = 8;
const PREC_INTERSECT = 9;
const PREC_RANGE = 10;

const OPERAND_START = new Set(['ref', 'name', 'func', 'open', 'structured']);

/** Parses a formula (without, or with, its leading `=`) into a syntax tree. Throws FormulaError. */
export function parseFormula(formula: string): FormulaAst {
	const all = tokenize(formula);
	const tokens: Token[] = [];
	const wsBefore: boolean[] = [];
	let sawWs = false;
	for (const token of all) {
		if (token.kind === 'ws') {
			sawWs = true;
			continue;
		}
		tokens.push(token);
		wsBefore.push(sawWs);
		sawWs = false;
	}
	if (tokens.length === 0) throw new FormulaError('Empty formula', 0);
	const cursor: Cursor = { tokens, wsBefore, pos: 0 };
	const ast = parseExpression(cursor, 0, false);
	const extra = cursor.tokens[cursor.pos];
	if (extra) throw new FormulaError(`Unexpected '${extra.text}'`, extra.start);
	return ast;
}

const peek = (c: Cursor): Token | undefined => c.tokens[c.pos];

function expect(c: Cursor, kind: Token['kind'], what: string): Token {
	const token = c.tokens[c.pos];
	if (!token || token.kind !== kind) {
		throw new FormulaError(`Expected ${what}`, token?.start ?? -1);
	}
	c.pos++;
	return token;
}

function parseExpression(c: Cursor, minPrec: number, inParens: boolean): FormulaAst {
	let left = parsePrefix(c, inParens);
	for (;;) {
		const token = peek(c);
		if (!token) return left;
		if (token.kind === 'op' && token.text === '%') {
			if (PREC_PERCENT < minPrec) return left;
			c.pos++;
			left = { type: 'percent', operand: left };
			continue;
		}
		if (
			c.wsBefore[c.pos] &&
			OPERAND_START.has(token.kind) &&
			PREC_INTERSECT >= minPrec &&
			isReferenceLike(left)
		) {
			const right = parseExpression(c, PREC_INTERSECT + 1, inParens);
			left = { type: 'binary', op: ' ', left, right };
			continue;
		}
		if (token.kind === 'comma' && inParens && PREC_UNION >= minPrec) {
			c.pos++;
			const right = parseExpression(c, PREC_UNION + 1, inParens);
			left = { type: 'binary', op: ',', left, right };
			continue;
		}
		if (token.kind === 'open' && (left.type === 'call' || left.type === 'invoke')) {
			c.pos++;
			left = { type: 'invoke', callee: left, args: parseArguments(c) };
			continue;
		}
		const binary = token.kind === 'op' ? BINARY[token.text] : undefined;
		if (!binary || binary.prec < minPrec) return left;
		c.pos++;
		const right = parseExpression(c, binary.prec + 1, inParens);
		left = { type: 'binary', op: binary.op, left, right };
	}
}

function isReferenceLike(node: FormulaAst): boolean {
	switch (node.type) {
		case 'ref':
		case 'name':
		case 'structured':
		case 'call':
			return true;
		case 'binary':
			return node.op === ':' || node.op === ' ' || node.op === ',';
		default:
			return false;
	}
}

function parsePrefix(c: Cursor, inParens: boolean): FormulaAst {
	const token = peek(c);
	if (!token) throw new FormulaError('Unexpected end of formula', -1);
	if (token.kind === 'op' && (token.text === '-' || token.text === '+')) {
		c.pos++;
		const operand = parseExpression(c, PREC_UNARY, inParens);
		return { type: 'unary', op: token.text, operand };
	}
	if (token.kind === 'op' && token.text === '@') {
		c.pos++;
		const operand = parseExpression(c, PREC_RANGE, inParens);
		return { type: 'unary', op: '@', operand };
	}
	c.pos++;
	switch (token.kind) {
		case 'number':
			return { type: 'number', value: token.value as number };
		case 'string':
			return { type: 'string', value: token.value as string };
		case 'bool':
			return { type: 'boolean', value: token.value as boolean };
		case 'error':
			return { type: 'error', code: token.value as ErrorCode };
		case 'ref': {
			const node: FormulaAst = { type: 'ref' };
			if (token.prefix) node.prefix = token.prefix;
			if (token.ref) node.ref = token.ref;
			if (token.spill) node.spill = true;
			return node;
		}
		case 'name': {
			const node: FormulaAst = { type: 'name', name: token.value as string };
			if (token.prefix) node.prefix = token.prefix;
			return node;
		}
		case 'structured':
			return {
				type: 'structured',
				ref: parseStructured(token.text, token.value as string),
				text: token.text,
			};
		case 'func': {
			expect(c, 'open', "'('");
			const rawName = token.value as string;
			return {
				type: 'call',
				name: normalizeFunctionName(rawName),
				rawName,
				args: parseArguments(c),
			};
		}
		case 'open': {
			const inner = parseExpression(c, 0, true);
			expect(c, 'close', "')'");
			return inner;
		}
		case 'lbrace':
			return parseArray(c);
		default:
			throw new FormulaError(`Unexpected '${token.text}'`, token.start);
	}
}

/** Arguments after `(` up to and including `)`; empty slots become `missing`. */
function parseArguments(c: Cursor): FormulaAst[] {
	const args: FormulaAst[] = [];
	if (peek(c)?.kind === 'close') {
		c.pos++;
		return args;
	}
	for (;;) {
		const token = peek(c);
		if (token?.kind === 'comma' || token?.kind === 'close') {
			args.push({ type: 'missing' });
		} else {
			args.push(parseExpression(c, 0, false));
		}
		const next = peek(c);
		if (next?.kind === 'comma') {
			c.pos++;
			continue;
		}
		expect(c, 'close', "',' or ')'");
		return args;
	}
}

function parseArrayItem(c: Cursor): ArrayConstant {
	let token = peek(c);
	let sign = 1;
	while (token?.kind === 'op' && (token.text === '-' || token.text === '+')) {
		if (token.text === '-') sign = -sign;
		c.pos++;
		token = peek(c);
	}
	if (!token) throw new FormulaError('Unterminated array constant', -1);
	c.pos++;
	if (token.kind === 'number') return sign * (token.value as number);
	if (sign !== 1) throw new FormulaError('Invalid array constant', token.start);
	if (token.kind === 'string') return token.value as string;
	if (token.kind === 'bool') return token.value as boolean;
	if (token.kind === 'error') return { error: token.value as ErrorCode };
	if (token.kind === 'ref' && !token.ref && !token.prefix) return { error: '#REF!' };
	throw new FormulaError(
		'Array constants may only hold numbers, text, logicals and errors',
		token.start,
	);
}

function parseArray(c: Cursor): FormulaAst {
	const rows: ArrayConstant[][] = [[]];
	for (;;) {
		const row = rows[rows.length - 1] ?? [];
		row.push(parseArrayItem(c));
		const token = peek(c);
		c.pos++;
		if (token?.kind === 'comma') continue;
		if (token?.kind === 'semicolon') {
			rows.push([]);
			continue;
		}
		if (token?.kind === 'rbrace') break;
		throw new FormulaError("Expected ',', ';' or '}'", token?.start ?? -1);
	}
	const width = rows[0]?.length ?? 0;
	if (rows.some((row) => row.length !== width)) {
		throw new FormulaError('Array constant rows must have the same length', -1);
	}
	return { type: 'array', rows };
}
