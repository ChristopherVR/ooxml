import { compileRegex, groupText, type Match, matches, type Program, search } from '../regex';
import { chargeText, checkArraySize, ERR, fail, Matrix, type Value } from '../values';
import { int, optNum, spec, str } from './helpers';
import type { FunctionSpec } from './types';

const C = 'Text';
const MAX_PATTERN_LENGTH = 1000;

/**
 * A pattern in JavaScript RegExp syntax with the `u` flag (`i` for case_sensitivity 1). Excel uses
 * PCRE2, so PCRE-only syntax (possessive quantifiers, recursion, inline comments) is `#VALUE!`.
 * The patterns run on this engine's own matcher, which tries each (instruction, position) state
 * at most once and charges every step to the recalculation's work budget, so no pattern can hang a
 * recalculation: one that would run too long is `#NUM!`.
 */
function compile(pattern: string, caseMode: number): Program {
	if (caseMode !== 0 && caseMode !== 1) fail(ERR.VALUE);
	if (pattern.length > MAX_PATTERN_LENGTH) fail(ERR.VALUE);
	try {
		return compileRegex(pattern, caseMode === 1);
	} catch {
		return fail(ERR.VALUE);
	}
}

const caseMode = (args: Value[], index: number): number => optNum(args, index, 0);

/** The replacement for one match, with `$$`, `$&`, `` $` ``, `$'`, `$n`, `$nn` and `$<name>`. */
function substitute(template: string, text: string, match: Match, program: Program): string {
	if (!template.includes('$')) return template;
	const groups = program.groupCount;
	let out = '';
	for (let i = 0; i < template.length; i++) {
		const ch = template[i] as string;
		const next = template[i + 1];
		if (ch !== '$' || next === undefined) {
			out += ch;
			continue;
		}
		if (next === '$' || next === '&' || next === '`' || next === "'") {
			i++;
			if (next === '$') out += '$';
			else if (next === '&') out += text.slice(match[0], match[1]);
			else if (next === '`') out += text.slice(0, match[0]);
			else out += text.slice(match[1]);
			continue;
		}
		if (next >= '0' && next <= '9') {
			let index = Number(next);
			let digits = 1;
			const second = template[i + 2] ?? '';
			const two = index * 10 + Number(second);
			if (second >= '0' && second <= '9' && two >= 1 && two <= groups) {
				index = two;
				digits = 2;
			}
			if (index >= 1 && index <= groups) {
				out += groupText(text, match, index) ?? '';
				i += digits;
			} else out += '$';
			continue;
		}
		const close = next === '<' && program.names.size ? template.indexOf('>', i + 2) : -1;
		if (close < 0) {
			out += '$';
			continue;
		}
		const indexes = program.names.get(template.slice(i + 2, close)) ?? [];
		const index = indexes.find((g) => groupText(text, match, g) !== undefined);
		out += index === undefined ? '' : (groupText(text, match, index) ?? '');
		i = close;
	}
	return out;
}

/** `text` with `picked` matches replaced, charging the text it builds. */
function replaceMatches(
	text: string,
	picked: Iterable<Match>,
	replacement: string,
	program: Program,
): string {
	let out = '';
	let last = 0;
	for (const match of picked) {
		out += chargeText(text.slice(last, match[0]));
		out += chargeText(substitute(replacement, text, match, program));
		last = match[1] as number;
	}
	return out + chargeText(text.slice(last));
}

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
		(args) => search(compile(str(args[1]), caseMode(args, 2)), str(args[0]), 0) !== null,
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
			const program = compile(str(args[1]), caseMode(args, 3));
			const text = str(args[0]);
			if (mode === 1) {
				const found: string[] = [];
				for (const match of matches(program, text)) {
					checkArraySize(found.length + 1, 1);
					found.push(text.slice(match[0], match[1]));
				}
				return found.length ? new Matrix(found.map((match) => [match])) : fail(ERR.NA);
			}
			const match = search(program, text, 0) ?? fail(ERR.NA);
			if (mode === 0) return text.slice(match[0], match[1]);
			if (program.groupCount === 0) fail(ERR.VALUE);
			const row = Array.from({ length: program.groupCount }, (_, g) =>
				groupText(text, match, g + 1),
			);
			return new Matrix([row.map((group) => group ?? '')]);
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
			const program = compile(str(args[1]), caseMode(args, 4));
			const replacement = str(args[2]);
			const occurrence = args.length > 3 && args[3] !== null ? int(args[3]) : 0;
			if (occurrence === 0)
				return replaceMatches(text, matches(program, text), replacement, program);
			let picked: Match | undefined;
			if (occurrence > 0) {
				let seen = 0;
				for (const match of matches(program, text))
					if (++seen === occurrence && (picked = match)) break;
			} else picked = [...matches(program, text)].at(occurrence);
			return picked ? replaceMatches(text, [picked], replacement, program) : text;
		},
	),
];
