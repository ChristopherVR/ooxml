import { checkArraySize, ERR, fail, Matrix, type Value } from '../values';
import { int, optNum, spec, str } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Text';

/**
 * A JavaScript regular expression for an Excel pattern. Excel uses PCRE2, so patterns that rely on
 * PCRE-only syntax (possessive quantifiers, recursion, inline comments) are `#VALUE!` here.
 */
function compile(pattern: string, caseMode: number, global: boolean): RegExp {
	if (caseMode !== 0 && caseMode !== 1) fail(ERR.VALUE);
	if (pattern.length > MAX_PATTERN_LENGTH || hasNestedQuantifier(pattern)) fail(ERR.VALUE);
	try {
		return new RegExp(pattern, `${global ? 'g' : ''}${caseMode === 1 ? 'i' : ''}u`);
	} catch {
		return fail(ERR.VALUE);
	}
}

const MAX_PATTERN_LENGTH = 1000;
const MAX_INPUT_LENGTH = 100_000;

/** The input is bounded too, so even a quadratic pattern stays fast. */
function limited(text: string): string {
	return text.length > MAX_INPUT_LENGTH ? fail(ERR.VALUE) : text;
}

const REPEAT = /^\{\d+,\d*\}/;

/**
 * Limitation: patterns run on the JavaScript backtracking matcher, so patterns that can backtrack
 * exponentially are refused with `#VALUE!` instead of hanging a recalculation. A pattern is
 * refused when it is longer than 1,000 characters, or when a group is repeated (`*`, `+`,
 * `{n,}`, `{n,m}`) and holds another quantifier or an alternation, such as `(a+)+`, `(a*)*` or
 * `(a|aa)+`. Excel (PCRE2) accepts these; rewrite them without the nesting.
 */
export function hasNestedQuantifier(pattern: string): boolean {
	// One entry per open group: whether it holds a quantifier, and whether it holds a `|`.
	const stack: { quantified: boolean; alternation: boolean }[] = [];
	let inClass = false;
	for (let i = 0; i < pattern.length; i++) {
		const ch = pattern[i];
		if (ch === '\\') {
			i++;
			continue;
		}
		if (inClass) {
			if (ch === ']') inClass = false;
			continue;
		}
		const top = stack[stack.length - 1];
		if (ch === '[') inClass = true;
		else if (ch === '(') stack.push({ quantified: false, alternation: false });
		else if (ch === '|') {
			if (top) top.alternation = true;
		} else if (ch === ')') {
			const group = stack.pop();
			const next = pattern[i + 1];
			const repeats =
				next === '*' || next === '+' || (next === '{' && REPEAT.test(pattern.slice(i + 1)));
			if (repeats && group && (group.quantified || group.alternation)) return true;
			const parent = stack[stack.length - 1];
			if (parent && (repeats || next === '?' || group?.quantified)) parent.quantified = true;
		} else if (top && (ch === '*' || ch === '+' || (ch === '{' && REPEAT.test(pattern.slice(i))))) {
			top.quantified = true;
		}
	}
	return false;
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
		(args) => compile(str(args[1]), caseMode(args, 2), false).test(limited(str(args[0]))),
	),
	spec(
		'REGEXEXTRACT',
		C,
		'REGEXEXTRACT(text, pattern, [return_mode], [case_sensitivity])',
		'Extracts the first match (0), all matches (1), or the first match’s capture groups (2).',
		2,
		4,
		(args) => {
			const mode = optNum(args, 2, 0);
			if (mode !== 0 && mode !== 1 && mode !== 2) fail(ERR.VALUE);
			const regex = compile(str(args[1]), caseMode(args, 3), mode === 1);
			const text = limited(str(args[0]));
			if (mode === 1) {
				const matches: string[] = [];
				// matchAll advances after zero-length matches, including Unicode code points.
				for (const match of text.matchAll(regex)) {
					checkArraySize(matches.length + 1, 1);
					matches.push(match[0]);
				}
				return matches.length ? new Matrix(matches.map((match) => [match])) : fail(ERR.NA);
			}
			const match = regex.exec(text);
			if (!match) fail(ERR.NA);
			if (mode === 0) return match[0];
			if (match.length === 1) fail(ERR.VALUE);
			return new Matrix([match.slice(1).map((group) => group ?? '')]);
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
			const text = limited(str(args[0]));
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
