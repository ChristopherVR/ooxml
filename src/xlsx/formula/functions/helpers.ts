import { check, compareScalars, finite, toBool, toNumber, toText } from '../coerce.js';
import type { CallContext } from '../context.js';
import { parseNumberText } from '../text-number.js';
import {
	ERR,
	ErrorSignal,
	fail,
	isError,
	isScalar,
	Matrix,
	type Scalar,
	type Value,
} from '../values.js';
import type { FunctionSpec, ParamKind } from './types.js';

/** A scalar argument (lifting already happened); anything else is `#VALUE!`. */
export function scalar(value: Value | undefined): Scalar {
	if (value === undefined) return null;
	if (!isScalar(value)) {
		if (value instanceof Matrix) return value.get(0, 0);
		fail(ERR.VALUE);
	}
	return value;
}

export const num = (value: Value | undefined): number => toNumber(scalar(value));
export const str = (value: Value | undefined): string => toText(scalar(value));
export const bool = (value: Value | undefined): boolean => toBool(scalar(value));
/** Truncates toward zero, like Excel's integer arguments. */
export const int = (value: Value | undefined): number => Math.trunc(num(value));

export const optNum = (args: Value[], i: number, fallback: number): number =>
	i < args.length ? num(args[i]) : fallback;
export const optBool = (args: Value[], i: number, fallback: boolean): boolean =>
	i < args.length ? bool(args[i]) : fallback;
export const optStr = (args: Value[], i: number, fallback: string): string =>
	i < args.length ? str(args[i]) : fallback;

/** Whether an optional argument was omitted (or left empty). */
export const omitted = (args: Value[], i: number): boolean => i >= args.length || args[i] === null;

interface CollectOptions {
	/** Count logicals and numeric text from references and arrays (the `A` functions). */
	includeAll?: boolean;
	/** Skip errors instead of throwing. */
	ignoreErrors?: boolean;
}

/**
 * The numbers of aggregate-function arguments with Excel's rules: direct arguments coerce text
 * and logicals; references and arrays contribute numbers only (unless `includeAll`).
 */
export function collectNumbers(
	ctx: CallContext,
	args: Value[],
	options: CollectOptions = {},
): number[] {
	const out: number[] = [];
	for (const arg of args) {
		ctx.forEach(arg, (value, kind) => {
			if (typeof value === 'number') {
				out.push(value);
				return;
			}
			if (isError(value)) {
				if (options.ignoreErrors) return;
				throw new ErrorSignal(value);
			}
			if (kind === 'direct') {
				if (value === null) out.push(0);
				else out.push(toNumber(value));
				return;
			}
			if (options.includeAll) {
				if (typeof value === 'boolean') out.push(value ? 1 : 0);
				else if (typeof value === 'string') out.push(0);
			}
		});
	}
	return out;
}

/** The numbers of the `A` functions (AVERAGEA, MAXA, STDEVA...): logicals and text count. */
export const all = (ctx: CallContext, args: Value[]): number[] =>
	collectNumbers(ctx, args, { includeAll: true });

/** `value` when `ok`, otherwise `#NUM!` (function domain checks). */
export const domain = (ok: boolean, value: number): number => (ok ? value : fail(ERR.NUM));

/** Every value of an argument in row-major order (refs read densely). */
export function flatValues(ctx: CallContext, value: Value): Scalar[] {
	return ctx.toMatrix(value).flat();
}

/** Builds a case-insensitive regular expression for `*`, `?` and `~` wildcards. */
export function wildcardRegex(pattern: string): RegExp {
	let source = '';
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i] ?? '';
		if (ch === '~' && i + 1 < pattern.length) {
			source += escapeRegex(pattern[i + 1] ?? '');
			i++;
		} else if (ch === '*') source += '[\\s\\S]*';
		else if (ch === '?') source += '[\\s\\S]';
		else source += escapeRegex(ch);
	}
	return new RegExp(`^${source}$`, 'i');
}

const escapeRegex = (text: string): string => text.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
export const hasWildcards = (text: string): boolean => /[*?~]/.test(text);

/** A predicate for COUNTIF-style criteria (`">5"`, `"<>x"`, `"a*"`, 7, TRUE, `""`). */
export function makeCriteria(criteria: Scalar): (value: Scalar) => boolean {
	if (isError(criteria)) return (value) => isError(value) && value.error === criteria.error;
	if (typeof criteria === 'number') {
		return (value) =>
			typeof value === 'number'
				? compareScalars(value, criteria) === 0
				: typeof value === 'string' && parseNumberText(value) === criteria;
	}
	if (typeof criteria === 'boolean') return (value) => value === criteria;
	if (criteria === null) return (value) => value === null || value === '';
	const match = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(criteria);
	const op = match?.[1] ?? '';
	const operandText = match?.[2] ?? '';
	const asNumber = parseNumberText(operandText);
	const upper = operandText.toUpperCase();
	const asBool = upper === 'TRUE' ? true : upper === 'FALSE' ? false : undefined;
	const errorOperand = /^#/.test(operandText) ? operandText.toUpperCase() : undefined;
	if (op === '' || op === '=' || op === '<>') {
		let test: (value: Scalar) => boolean;
		if (operandText === '') {
			test = (value) => value === null || value === '';
			if (op === '') return test;
		} else if (asNumber !== undefined) {
			test = (value) =>
				(typeof value === 'number' && compareScalars(value, asNumber) === 0) ||
				(typeof value === 'string' && parseNumberText(value) === asNumber);
		} else if (asBool !== undefined) {
			test = (value) => value === asBool;
		} else if (errorOperand) {
			test = (value) => isError(value) && value.error === errorOperand;
		} else if (hasWildcards(operandText)) {
			const regex = wildcardRegex(operandText);
			test = (value) => typeof value === 'string' && regex.test(value);
		} else {
			const lower = operandText.toLowerCase();
			test = (value) => typeof value === 'string' && value.toLowerCase() === lower;
		}
		if (op === '<>') {
			if (operandText === '') return (value) => value !== null && value !== '';
			return (value) => !test(value);
		}
		return test;
	}
	const sign = (c: number): boolean =>
		op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
	if (asNumber !== undefined) {
		return (value) => typeof value === 'number' && sign(compareScalars(value, asNumber));
	}
	return (value) => typeof value === 'string' && sign(compareScalars(value, operandText));
}

/** A FunctionSpec in a compact form. */
export function spec(
	name: string,
	category: string,
	syntax: string,
	description: string,
	minArgs: number,
	maxArgs: number,
	fn: (args: Value[], ctx: CallContext) => Value,
	params?: readonly ParamKind[],
	volatile?: boolean,
): FunctionSpec {
	const out: FunctionSpec = { name, category, syntax, description, minArgs, maxArgs, fn };
	if (params) out.params = params;
	if (volatile) out.volatile = true;
	return out;
}

/** A numeric function of fixed arity whose arguments all coerce to numbers. */
export function numeric(
	name: string,
	category: string,
	syntax: string,
	description: string,
	minArgs: number,
	maxArgs: number,
	fn: (...values: number[]) => number | Scalar,
	defaults: readonly number[] = [],
): FunctionSpec {
	return spec(name, category, syntax, description, minArgs, maxArgs, (args) => {
		const values: number[] = [];
		for (let i = 0; i < maxArgs; i++) {
			if (i < args.length) values.push(num(args[i]));
			else if (defaults[i] !== undefined) values.push(defaults[i] as number);
		}
		const result = fn(...values);
		return typeof result === 'number' ? finite(result) : result;
	});
}

export { check, finite, toBool, toNumber, toText };
