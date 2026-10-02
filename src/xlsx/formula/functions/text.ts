import { toText } from '../coerce.js';
import type { CallContext } from '../context.js';
import { ERR, fail, isError, type Value } from '../values.js';
import { bool, num, optNum, spec, str, wildcardRegex } from './helpers.js';
import type { FunctionSpec } from './types.js';
import { TEXT_CONVERT } from './text-convert.js';

const C = 'Text';
const MAX_TEXT = 32_767;

const limit = (text: string): string => (text.length > MAX_TEXT ? fail(ERR.VALUE) : text);

function find(args: Value[], insensitive: boolean): number {
	const needle = str(args[0]);
	const haystack = str(args[1]);
	const start = optNum(args, 2, 1);
	if (start < 1 || start > haystack.length + 1) fail(ERR.VALUE);
	const from = Math.trunc(start) - 1;
	if (needle === '') return from + 1;
	if (!insensitive) {
		const at = haystack.indexOf(needle, from);
		return at < 0 ? fail(ERR.VALUE) : at + 1;
	}
	const regex = wildcardRegex(needle);
	const pattern = new RegExp(regex.source.slice(1, -1), 'i');
	const match = pattern.exec(haystack.slice(from));
	return match ? from + match.index + 1 : fail(ERR.VALUE);
}

function proper(text: string): string {
	let out = '';
	let previousLetter = false;
	for (const ch of text) {
		const isLetter = ch.toLowerCase() !== ch.toUpperCase();
		out += isLetter ? (previousLetter ? ch.toLowerCase() : ch.toUpperCase()) : ch;
		previousLetter = isLetter;
	}
	return out;
}

/** Every value of the arguments as text, blanks included (ranges are read densely). */
function joinValues(ctx: CallContext, args: Value[]): string[] {
	const out: string[] = [];
	for (const arg of args) {
		for (const v of ctx.toMatrix(arg).flat()) {
			if (isError(v)) fail(v);
			out.push(toText(v));
		}
	}
	return out;
}

export const TEXT_FUNCTIONS: FunctionSpec[] = [
	...TEXT_CONVERT,
	spec('CONCATENATE', C, 'CONCATENATE(text1, [text2], ...)', 'Joins text items.', 1, 255, (args) =>
		limit(args.map((a) => str(a)).join('')),
	),
	spec(
		'CONCAT',
		C,
		'CONCAT(text1, [text2], ...)',
		'Joins text from items and ranges.',
		1,
		255,
		(args, ctx) => limit(joinValues(ctx, args).join('')),
		['any'],
	),
	spec(
		'TEXTJOIN',
		C,
		'TEXTJOIN(delimiter, ignore_empty, text1, ...)',
		'Joins text with a delimiter.',
		3,
		255,
		(args, ctx) => {
			const delimiters = ctx
				.toMatrix(args[0] ?? null)
				.flat()
				.map((d) => toText(d));
			const ignoreEmpty = bool(ctx.toScalar(args[1] ?? null));
			const parts = joinValues(ctx, args.slice(2)).filter((p) => !ignoreEmpty || p !== '');
			let out = '';
			parts.forEach((p, i) => {
				if (i > 0) out += delimiters.length ? delimiters[(i - 1) % delimiters.length] : '';
				out += p;
			});
			return limit(out);
		},
		['any'],
	),
	...(['LEFT', 'LEFTB'] as const).map((name) =>
		spec(name, C, `${name}(text, [num_chars])`, 'The first characters of a text.', 1, 2, (args) => {
			const n = optNum(args, 1, 1);
			if (n < 0) fail(ERR.VALUE);
			return str(args[0]).slice(0, Math.trunc(n));
		}),
	),
	...(['RIGHT', 'RIGHTB'] as const).map((name) =>
		spec(name, C, `${name}(text, [num_chars])`, 'The last characters of a text.', 1, 2, (args) => {
			const n = Math.trunc(optNum(args, 1, 1));
			if (n < 0) fail(ERR.VALUE);
			const text = str(args[0]);
			return n === 0 ? '' : text.slice(Math.max(0, text.length - n));
		}),
	),
	...(['MID', 'MIDB'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(text, start_num, num_chars)`,
			'Characters from the middle of a text.',
			3,
			3,
			(args) => {
				const start = Math.trunc(num(args[1]));
				const n = Math.trunc(num(args[2]));
				if (start < 1 || n < 0) fail(ERR.VALUE);
				return str(args[0]).substr(start - 1, n);
			},
		),
	),
	...(['LEN', 'LENB'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(text)`,
			'The number of characters in a text.',
			1,
			1,
			(args) => str(args[0]).length,
		),
	),
	spec('LOWER', C, 'LOWER(text)', 'Converts text to lower case.', 1, 1, (args) =>
		str(args[0]).toLowerCase(),
	),
	spec('UPPER', C, 'UPPER(text)', 'Converts text to upper case.', 1, 1, (args) =>
		str(args[0]).toUpperCase(),
	),
	spec('PROPER', C, 'PROPER(text)', 'Capitalizes the first letter of each word.', 1, 1, (args) =>
		proper(str(args[0])),
	),
	spec('TRIM', C, 'TRIM(text)', 'Removes extra spaces.', 1, 1, (args) =>
		str(args[0]).replace(/ +/g, ' ').replace(/^ | $/g, ''),
	),
	spec('CLEAN', C, 'CLEAN(text)', 'Removes non-printable characters.', 1, 1, (args) =>
		// eslint-disable-next-line no-control-regex
		str(args[0]).replace(/[\x00-\x1f]/g, ''),
	),
	spec(
		'SUBSTITUTE',
		C,
		'SUBSTITUTE(text, old_text, new_text, [instance_num])',
		'Replaces occurrences of a text.',
		3,
		4,
		(args) => {
			const text = str(args[0]);
			const oldText = str(args[1]);
			const newText = str(args[2]);
			if (oldText === '') return text;
			if (args.length < 4) return limit(text.split(oldText).join(newText));
			const instance = Math.trunc(num(args[3]));
			if (instance < 1) fail(ERR.VALUE);
			let at = -1;
			for (let i = 0; i < instance; i++) {
				at = text.indexOf(oldText, at + 1);
				if (at < 0) return text;
			}
			return limit(text.slice(0, at) + newText + text.slice(at + oldText.length));
		},
	),
	...(['REPLACE', 'REPLACEB'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(old_text, start_num, num_chars, new_text)`,
			'Replaces part of a text.',
			4,
			4,
			(args) => {
				const text = str(args[0]);
				const start = Math.trunc(num(args[1]));
				const n = Math.trunc(num(args[2]));
				if (start < 1 || n < 0) fail(ERR.VALUE);
				return limit(text.slice(0, start - 1) + str(args[3]) + text.slice(start - 1 + n));
			},
		),
	),
	...(['FIND', 'FINDB'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(find_text, within_text, [start_num])`,
			'The position of a text (case-sensitive).',
			2,
			3,
			(args) => find(args, false),
		),
	),
	...(['SEARCH', 'SEARCHB'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(find_text, within_text, [start_num])`,
			'The position of a text (wildcards, any case).',
			2,
			3,
			(args) => find(args, true),
		),
	),
	spec('REPT', C, 'REPT(text, number_times)', 'Repeats a text.', 2, 2, (args) => {
		const n = Math.trunc(num(args[1]));
		const text = str(args[0]);
		if (n < 0 || text.length * n > MAX_TEXT) fail(ERR.VALUE);
		return text.repeat(n);
	}),
	spec(
		'EXACT',
		C,
		'EXACT(text1, text2)',
		'Whether two texts are identical (case-sensitive).',
		2,
		2,
		(args) => str(args[0]) === str(args[1]),
	),
];
