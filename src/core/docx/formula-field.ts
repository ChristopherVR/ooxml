// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word `=` formula field (ECMA-376 Part 1, 17.16.3): arithmetic, comparisons, and built-in
// scalar/list functions without table-cell or bookmark references.
import { applyFieldFormat, parseFieldInstruction } from './field-format';

type Token =
	| { kind: 'num'; value: number }
	| { kind: 'id'; value: string }
	| { kind: 'op'; value: string }
	| { kind: 'punct'; value: '(' | ')' | ',' };

const MAX_DEPTH = 64;

function tokenize(expr: string): Token[] | undefined {
	const tokens: Token[] = [];
	let i = 0;
	while (i < expr.length) {
		const ch = expr[i]!;
		if (/\s/.test(ch)) {
			i++;
			continue;
		}
		if (ch === '(' || ch === ')' || ch === ',') {
			tokens.push({ kind: 'punct', value: ch });
			i++;
			continue;
		}
		if (
			(ch === '<' || ch === '>') &&
			(expr[i + 1] === '=' || (ch === '<' && expr[i + 1] === '>'))
		) {
			tokens.push({ kind: 'op', value: ch + expr[i + 1]! });
			i += 2;
			continue;
		}
		if ('+-*/^%=<>'.includes(ch)) {
			tokens.push({ kind: 'op', value: ch });
			i++;
			continue;
		}
		if (/[0-9.]/.test(ch)) {
			const start = i;
			let dot = ch === '.';
			i++;
			while (i < expr.length) {
				const c = expr[i]!;
				if (c >= '0' && c <= '9') i++;
				else if (c === '.' && !dot) {
					dot = true;
					i++;
				} else break;
			}
			const n = Number(expr.slice(start, i));
			if (!Number.isFinite(n)) return undefined;
			tokens.push({ kind: 'num', value: n });
			continue;
		}
		if (/[A-Za-z_]/.test(ch)) {
			const start = i;
			i++;
			while (i < expr.length && /[A-Za-z0-9_.]/.test(expr[i]!)) i++;
			tokens.push({ kind: 'id', value: expr.slice(start, i).toUpperCase() });
			continue;
		}
		return undefined;
	}
	return tokens;
}

class ZeroDivideError extends Error {}

function evalFunction(name: string, args: number[]): number | undefined {
	switch (name) {
		case 'ABS':
			return args.length === 1 ? Math.abs(args[0]!) : undefined;
		case 'INT':
			return args.length === 1 ? Math.floor(args[0]!) : undefined;
		case 'SIGN':
			return args.length === 1 ? Math.sign(args[0]!) : undefined;
		case 'NOT':
			return args.length === 1 ? (args[0] === 0 ? 1 : 0) : undefined;
		case 'MOD': {
			if (args.length !== 2) return undefined;
			const [a, b] = [args[0]!, args[1]!];
			if (b === 0) throw new ZeroDivideError();
			return a - b * Math.floor(a / b);
		}
		case 'ROUND': {
			if (args.length !== 2) return undefined;
			const [value, digits] = [args[0]!, Math.trunc(args[1]!)];
			const factor = 10 ** digits;
			return (Math.sign(value) * Math.round(Math.abs(value) * factor + Number.EPSILON)) / factor;
		}
		case 'IF':
			return args.length === 3 ? (args[0] !== 0 ? args[1]! : args[2]!) : undefined;
		case 'AND':
			return args.length >= 1 ? (args.every((v) => v !== 0) ? 1 : 0) : undefined;
		case 'OR':
			return args.length >= 1 ? (args.some((v) => v !== 0) ? 1 : 0) : undefined;
		case 'SUM':
			return args.reduce((acc, v) => acc + v, 0);
		case 'PRODUCT':
			return args.length >= 1 ? args.reduce((acc, v) => acc * v, 1) : 0;
		case 'AVERAGE':
			if (args.length === 0) throw new ZeroDivideError();
			return args.reduce((acc, v) => acc + v, 0) / args.length;
		case 'COUNT':
			return args.length;
		case 'MIN':
			return args.length >= 1 ? Math.min(...args) : undefined;
		case 'MAX':
			return args.length >= 1 ? Math.max(...args) : undefined;
		default:
			return undefined;
	}
}

function evaluateTokens(tokens: Token[]): number | undefined {
	let pos = 0;
	let depth = 0;
	const peek = (): Token | undefined => tokens[pos];
	const next = (): Token | undefined => tokens[pos++];

	const parseComparison = (): number | undefined => {
		let left = parseAdditive();
		if (left === undefined) return undefined;
		const t = peek();
		if (t?.kind === 'op' && ['=', '<>', '<', '<=', '>', '>='].includes(t.value)) {
			next();
			const right = parseAdditive();
			if (right === undefined) return undefined;
			switch (t.value) {
				case '=':
					left = left === right ? 1 : 0;
					break;
				case '<>':
					left = left !== right ? 1 : 0;
					break;
				case '<':
					left = left < right ? 1 : 0;
					break;
				case '<=':
					left = left <= right ? 1 : 0;
					break;
				case '>':
					left = left > right ? 1 : 0;
					break;
				case '>=':
					left = left >= right ? 1 : 0;
					break;
			}
		}
		return left;
	};

	const parseAdditive = (): number | undefined => {
		let left = parseMultiplicative();
		if (left === undefined) return undefined;
		while (peek()?.kind === 'op' && (peek()!.value === '+' || peek()!.value === '-')) {
			const op = next()!.value;
			const right = parseMultiplicative();
			if (right === undefined) return undefined;
			left = op === '+' ? left + right : left - right;
		}
		return left;
	};

	const parseMultiplicative = (): number | undefined => {
		let left = parsePower();
		if (left === undefined) return undefined;
		while (peek()?.kind === 'op' && (peek()!.value === '*' || peek()!.value === '/')) {
			const op = next()!.value;
			const right = parsePower();
			if (right === undefined) return undefined;
			if (op === '/') {
				if (right === 0) throw new ZeroDivideError();
				left = left / right;
			} else {
				left = left * right;
			}
		}
		return left;
	};

	const parsePower = (): number | undefined => {
		const base = parseUnary();
		if (base === undefined) return undefined;
		if (peek()?.kind === 'op' && peek()!.value === '^') {
			next();
			const exp = parsePower();
			if (exp === undefined) return undefined;
			const result = base ** exp;
			return Number.isFinite(result) ? result : undefined;
		}
		return base;
	};

	const parseUnary = (): number | undefined => {
		const t = peek();
		if (t?.kind === 'op' && (t.value === '+' || t.value === '-')) {
			next();
			const v = parseUnary();
			if (v === undefined) return undefined;
			return t.value === '-' ? -v : v;
		}
		let v = parsePrimary();
		if (v === undefined) return undefined;
		while (peek()?.kind === 'op' && peek()!.value === '%') {
			next();
			v = v / 100;
		}
		return v;
	};

	const parsePrimary = (): number | undefined => {
		if (++depth > MAX_DEPTH) return undefined;
		try {
			const t = next();
			if (!t) return undefined;
			if (t.kind === 'num') return t.value;
			if (t.kind === 'punct' && t.value === '(') {
				const inner = parseComparison();
				const close = next();
				if (inner === undefined || close?.kind !== 'punct' || close.value !== ')') return undefined;
				return inner;
			}
			if (t.kind === 'id') {
				if (t.value === 'TRUE') return 1;
				if (t.value === 'FALSE') return 0;
				if (peek()?.kind !== 'punct' || peek()!.value !== '(') return undefined;
				next();
				const args: number[] = [];
				if (peek()?.kind !== 'punct' || peek()!.value !== ')') {
					while (true) {
						const arg = parseComparison();
						if (arg === undefined) return undefined;
						args.push(arg);
						if (peek()?.kind === 'punct' && peek()!.value === ',') {
							next();
							continue;
						}
						break;
					}
				}
				const close = next();
				if (close?.kind !== 'punct' || close.value !== ')') return undefined;
				return evalFunction(t.value, args);
			}
			return undefined;
		} finally {
			depth--;
		}
	};

	const result = parseComparison();
	if (result === undefined || pos !== tokens.length || !Number.isFinite(result)) return undefined;
	return result;
}

/** Splits `= expr [\* ...]` into its expression text and trailing field switches. */
function splitFormulaAndSwitches(instr: string): { expr: string; switches: string } | undefined {
	const trimmed = instr.trim();
	if (!trimmed.startsWith('=')) return undefined;
	const body = trimmed.slice(1);
	const switchIndex = body.search(/\\[A-Za-z*@#]/);
	if (switchIndex < 0) return { expr: body.trim(), switches: '' };
	return {
		expr: body.slice(0, switchIndex).trim(),
		switches: body.slice(switchIndex).trim(),
	};
}

export const ZERO_DIVIDE_ERROR = '!Zero Divide';

/**
 * Evaluates a Word `=` formula field. Returns undefined for expressions that reference table
 * cells, bookmarks, or unsupported switches (`\#`), so those fields keep their saved result.
 */
export function formulaFieldResult(instr: string): string | undefined {
	const parts = splitFormulaAndSwitches(instr);
	if (!parts || !parts.expr) return undefined;
	const format = parts.switches ? parseFieldInstruction(`= ${parts.switches}`)?.format : undefined;
	if (parts.switches && !format) return undefined;
	const tokens = tokenize(parts.expr);
	if (!tokens || tokens.length === 0) return undefined;
	try {
		const value = evaluateTokens(tokens);
		if (value === undefined) return undefined;
		const normalized = Object.is(value, -0) ? 0 : Math.round(value * 1e12) / 1e12;
		const text = String(normalized);
		return format ? applyFieldFormat(text, format) : text;
	} catch (error) {
		if (error instanceof ZeroDivideError) return ZERO_DIVIDE_ERROR;
		return undefined;
	}
}
