import { compile, type Program } from './compile';
import { codePointAt } from './charset';
import { parse } from './parse';
import { type Match, search } from './vm';

export type { Program } from './compile';
export type { Match } from './vm';
export { search } from './vm';

const cache = new Map<string, Program>();
const CACHE_SIZE = 64;

/**
 * Compiles a JavaScript pattern (`u` flag, and `i` when `ignoreCase`) for {@link search}. Throws
 * on a pattern the platform's RegExp or this parser rejects. Programs are cached by pattern.
 */
export function compileRegex(pattern: string, ignoreCase: boolean): Program {
	const key = `${ignoreCase ? 'i' : '-'}${pattern}`;
	const hit = cache.get(key);
	if (hit) return hit;
	// The platform validates the syntax; it is only constructed, never run.
	RegExp(pattern, ignoreCase ? 'ui' : 'u');
	const program = compile(parse(pattern, ignoreCase));
	if (cache.size >= CACHE_SIZE) cache.delete(cache.keys().next().value ?? '');
	cache.set(key, program);
	return program;
}

/** Every match from the start, as `String.prototype.matchAll` finds them with the `u` flag. */
export function* matches(program: Program, text: string): Generator<Match> {
	let from = 0;
	while (from <= text.length) {
		const match = search(program, text, from);
		if (!match) return;
		yield match;
		const start = match[0] as number;
		const end = match[1] as number;
		// An empty match moves on by one code point.
		from = end > start ? end : end + (end < text.length && codePointAt(text, end) > 0xffff ? 2 : 1);
	}
}

/** The text of group `g` of a match, or undefined when it did not take part. */
export function groupText(text: string, match: Match, g: number): string | undefined {
	const start = match[2 * g] ?? -1;
	const end = match[2 * g + 1] ?? -1;
	return start >= 0 && end >= 0 ? text.slice(start, end) : undefined;
}
