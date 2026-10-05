import { ERR, fail, type Value } from '../values.js';
import { int, optNum, spec, str } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Text';

/**
 * A JavaScript regular expression for an Excel pattern. Excel uses PCRE2, so patterns that rely on
 * PCRE-only syntax (possessive quantifiers, recursion, inline comments) are `#VALUE!` here.
 */
function compile(pattern: string, caseMode: number, global: boolean): RegExp {
	if (caseMode !== 0 && caseMode !== 1) fail(ERR.VALUE);
	try {
		return new RegExp(pattern, `${global ? 'g' : ''}${caseMode === 1 ? 'i' : ''}u`);
	} catch {
		return fail(ERR.VALUE);
	}
}

const caseMode = (args: Value[], index: number): number => optNum(args, index, 0);

/** RFC 3986 unreserved characters stay; everything else is percent-encoded as UTF-8. */
export function encodeUrl(text: string): string {
	return encodeURIComponent(text).replace(
		/[!'()*]/g,
		(c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`,
	);
}

export const TEXT_REGEX_FUNCTIONS: FunctionSpec[] = [
	spec('ENCODEURL', C, 'ENCODEURL(text)', 'Percent-encodes text for use in a URL.', 1, 1, (args) =>
		encodeUrl(str(args[0])),
	),
	spec(
		'REGEXTEST',
		C,
		'REGEXTEST(text, pattern, [case_sensitivity])',
		'Whether any part of the text matches the pattern.',
		2,
		3,
		(args) => compile(str(args[1]), caseMode(args, 2), false).test(str(args[0])),
	),
	spec(
		'REGEXEXTRACT',
		C,
		'REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])',
		'The first part of the text that matches the pattern.',
		2,
		4,
		(args) => {
			// Only the first-match mode is implemented; the all-matches and capture-group modes
			// return arrays whose layout is not confirmed against Excel.
			if (optNum(args, 2, 0) !== 0) fail(ERR.VALUE);
			const match = compile(str(args[1]), caseMode(args, 3), false).exec(str(args[0]));
			return match ? match[0] : fail(ERR.NA);
		},
	),
	spec(
		'REGEXREPLACE',
		C,
		'REGEXREPLACE(text, pattern, replacement, [occurrence], [case_sensitivity])',
		'Replaces the parts of the text that match the pattern.',
		3,
		5,
		(args) => {
			const text = str(args[0]);
			const regex = compile(str(args[1]), caseMode(args, 4), true);
			const replacement = str(args[2]);
			const occurrence = args.length > 3 && args[3] !== null ? int(args[3]) : 0;
			if (occurrence === 0) return text.replace(regex, replacement);
			const matches = [...text.matchAll(regex)];
			const picked = occurrence > 0 ? matches[occurrence - 1] : matches.at(occurrence);
			if (!picked || picked.index === undefined) return text;
			return (
				text.slice(0, picked.index) +
				picked[0].replace(new RegExp(regex.source, regex.flags.replace('g', '')), replacement) +
				text.slice(picked.index + picked[0].length)
			);
		},
	),
];
