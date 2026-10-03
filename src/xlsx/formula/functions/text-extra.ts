import { toText } from '../coerce.js';
import type { CallContext } from '../context.js';
import { ERR, fail, isError, Matrix, type Scalar, type Value } from '../values.js';
import { optBool, optNum, scalar, spec, str } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Text';

function delimitersOf(ctx: CallContext, value: Value | undefined): string[] {
	if (value === undefined || value === null) return [];
	return ctx
		.toMatrix(value)
		.flat()
		.map((v) => {
			if (isError(v)) fail(v);
			return toText(v);
		});
}

/** Positions of every delimiter occurrence, left to right, non-overlapping. */
function occurrences(
	text: string,
	delimiters: string[],
	insensitive: boolean,
): { at: number; length: number }[] {
	const hay = insensitive ? text.toLowerCase() : text;
	const needles = delimiters.map((d) => (insensitive ? d.toLowerCase() : d));
	const out: { at: number; length: number }[] = [];
	let i = 0;
	while (i <= hay.length) {
		let best: { at: number; length: number } | undefined;
		for (const n of needles) {
			if (n === '') continue;
			const at = hay.indexOf(n, i);
			if (at >= 0 && (!best || at < best.at || (at === best.at && n.length > best.length))) {
				best = { at, length: n.length };
			}
		}
		if (!best) break;
		out.push(best);
		i = best.at + best.length;
	}
	return out;
}

function textAround(args: Value[], ctx: CallContext, before: boolean): Value {
	const text = str(args[0]);
	const delimiters = delimitersOf(ctx, args[1]);
	const instance = Math.trunc(optNum(args, 2, 1));
	const insensitive = optNum(args, 3, 0) === 1;
	const matchEnd = optBool(args, 4, false);
	const notFound = args.length > 5 ? (scalar(args[5]) ?? ERR.NA) : ERR.NA;
	if (instance === 0 || Math.abs(instance) > text.length + 1) fail(ERR.VALUE);
	if (delimiters.length === 0 || delimiters.every((d) => d === '')) {
		return before ? (instance > 0 ? '' : text) : instance > 0 ? text : '';
	}
	const found = occurrences(text, delimiters, insensitive);
	if (matchEnd) {
		if (instance > 0) found.push({ at: text.length, length: 0 });
		else found.unshift({ at: 0, length: 0 });
	}
	const index = instance > 0 ? instance - 1 : found.length + instance;
	const hit = found[index];
	if (!hit || index < 0) return notFound;
	return before ? text.slice(0, hit.at) : text.slice(hit.at + hit.length);
}

function splitText(
	text: string,
	delimiters: string[],
	insensitive: boolean,
	ignoreEmpty: boolean,
): string[] {
	if (delimiters.length === 0) return [text];
	const parts: string[] = [];
	let last = 0;
	for (const hit of occurrences(text, delimiters, insensitive)) {
		parts.push(text.slice(last, hit.at));
		last = hit.at + hit.length;
	}
	parts.push(text.slice(last));
	return ignoreEmpty ? parts.filter((p) => p !== '') : parts;
}

function valueToText(value: Scalar, strict: boolean): string {
	if (isError(value)) return value.error;
	if (typeof value === 'string') return strict ? `"${value.replace(/"/g, '""')}"` : value;
	return toText(value);
}

export const TEXT_EXTRA_FUNCTIONS: FunctionSpec[] = [
	spec(
		'TEXTBEFORE',
		C,
		'TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])',
		'The text before a delimiter.',
		2,
		6,
		(args, ctx) => textAround(args, ctx, true),
		['value', 'any', 'value'],
	),
	spec(
		'TEXTAFTER',
		C,
		'TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])',
		'The text after a delimiter.',
		2,
		6,
		(args, ctx) => textAround(args, ctx, false),
		['value', 'any', 'value'],
	),
	spec(
		'TEXTSPLIT',
		C,
		'TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])',
		'Splits text into an array.',
		2,
		6,
		(args, ctx) => {
			const text = str(ctx.toScalar(args[0] ?? null));
			const colDelims = delimitersOf(ctx, args[1]);
			const rowDelims = delimitersOf(ctx, args[2]);
			const ignoreEmpty =
				args.length > 3 && args[3] !== null && Boolean(scalar(ctx.toScalar(args[3] ?? null)));
			const insensitive = args.length > 4 && optNum([ctx.toScalar(args[4] ?? null)], 0, 0) === 1;
			const pad: Scalar = args.length > 5 ? ctx.toScalar(args[5] ?? null) : ERR.NA;
			if (colDelims.length === 0 && rowDelims.length === 0) fail(ERR.VALUE);
			const rows = splitText(text, rowDelims, insensitive, ignoreEmpty).map((line) =>
				splitText(line, colDelims, insensitive, ignoreEmpty),
			);
			const width = rows.reduce((w, r) => Math.max(w, r.length), 0);
			return new Matrix(rows.map((r) => Array.from({ length: width }, (_, i) => r[i] ?? pad)));
		},
		['any'],
	),
	spec('VALUETOTEXT', C, 'VALUETOTEXT(value, [format])', 'Text for any value.', 1, 2, (args) =>
		valueToText(scalar(args[0]), optNum(args, 1, 0) === 1),
	),
	spec(
		'ARRAYTOTEXT',
		C,
		'ARRAYTOTEXT(array, [format])',
		'Text for an array.',
		1,
		2,
		(args, ctx) => {
			const strict = args.length > 1 && optNum([ctx.toScalar(args[1] ?? null)], 0, 0) === 1;
			const m = ctx.toMatrix(args[0] ?? null);
			if (!strict)
				return m
					.flat()
					.map((v) => valueToText(v, false))
					.join(', ');
			return `{${m.data.map((row) => row.map((v) => valueToText(v, true)).join(',')).join(';')}}`;
		},
		['any'],
	),
];
