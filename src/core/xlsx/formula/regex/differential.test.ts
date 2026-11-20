import { afterEach, describe, expect, it } from 'vitest';
import { resetCalcBudget } from '../values';
import { TEXT_REGEX_FUNCTIONS } from '../functions/text-regex';
import { compileRegex, groupText, type Match, matches, search } from '.';
import { tuneMemo } from './vm';
import { compile } from './compile';
import { parse } from './parse';

/**
 * Differential test: random patterns from the supported grammar, on random inputs, must give the
 * same match, index and captures as the platform's RegExp (which runs them with backtracking), and
 * each case must finish quickly.
 */

function random(seed: number) {
	let state = seed;
	const next = () => {
		state = (state * 48271) % 2147483647;
		return state / 2147483647;
	};
	const int = (n: number) => Math.floor(next() * n);
	const pick = <T>(items: readonly T[]): T => items[int(items.length)] as T;
	return { next, int, pick };
}

const ATOMS = [
	'a',
	'b',
	'c',
	'A',
	'.',
	'\\d',
	'\\w',
	'\\W',
	'\\s',
	'[ab]',
	'[^a]',
	'[a-c1]',
	'\\.',
	'😀',
	'ſ',
];
const ASSERTIONS = ['^', '$', '\\b', '\\B'];
const QUANTIFIERS = ['*', '+', '?', '{2}', '{1,}', '{0,2}', '{1,3}', '{2,4}'];

/** A random pattern; `groups` counts the capture groups so far, for backreferences. */
function pattern(rng: ReturnType<typeof random>, depth: number, groups: { count: number }): string {
	const term = (): string => {
		const roll = rng.int(depth > 2 ? 10 : 16);
		if (roll < 6) return rng.pick(ATOMS);
		if (roll < 7) return rng.pick(ASSERTIONS);
		if (roll < 8 && groups.count > 0) return `\\${1 + rng.int(groups.count)}`;
		if (roll < 10) return rng.pick(ATOMS);
		const kind = rng.int(9);
		if (kind < 3) {
			groups.count++;
			return `(${pattern(rng, depth + 1, groups)})`;
		}
		if (kind < 4) {
			groups.count++;
			return `(?<n${groups.count}>${pattern(rng, depth + 1, groups)})`;
		}
		if (kind < 7) return `(?:${pattern(rng, depth + 1, groups)})`;
		const look = rng.pick(['(?=', '(?!', '(?<=', '(?<!']);
		return `${look}${pattern(rng, depth + 1, groups)})`;
	};
	const sequence = (): string => {
		let out = '';
		for (let i = 0, n = 1 + rng.int(4); i < n; i++) {
			const t = term();
			const quantifiable = !ASSERTIONS.includes(t) && !/^\(\?<?[=!]/.test(t);
			out += t;
			if (quantifiable && rng.next() < 0.35)
				out += rng.pick(QUANTIFIERS) + (rng.next() < 0.3 ? '?' : '');
		}
		return out;
	};
	let out = sequence();
	while (rng.next() < 0.25) out += `|${sequence()}`;
	return out;
}

const ALPHABET = ['a', 'a', 'b', 'c', 'A', '1', ' ', '.', '_', '\n', '😀', 'ſ'];

function input(rng: ReturnType<typeof random>): string {
	let out = '';
	for (let i = 0, n = rng.int(12); i < n; i++) out += rng.pick(ALPHABET);
	return out;
}

/** The platform's answer in our shape: start, end and each group's text (undefined when unset). */
function expected(match: RegExpExecArray | null) {
	if (!match) return null;
	return { index: match.index, end: match.index + match[0].length, groups: match.slice(1) };
}

function actual(text: string, match: Match | null, groupCount: number) {
	if (!match) return null;
	const groups = Array.from({ length: groupCount }, (_, g) => groupText(text, match, g + 1));
	return { index: match[0], end: match[1], groups };
}

const isHigh = (c: number) => c >= 0xd800 && c < 0xdc00;
const isLow = (c: number) => c >= 0xdc00 && c <= 0xdfff;
/** V8 (unlike the specification and JavaScriptCore) can start a match inside a surrogate pair. */
const insidePair = (text: string, at: number) =>
	isHigh(text.charCodeAt(at - 1)) && isLow(text.charCodeAt(at));

const REPLACE = TEXT_REGEX_FUNCTIONS.find((f) => f.name === 'REGEXREPLACE')?.fn;
/** The memo from the first step, from a random step, and as searches normally start it. */
const MEMO = [
	(): number => 0,
	(rng: ReturnType<typeof random>): number => rng.int(40),
	null,
] as const;
const TEMPLATES = ['<$&>', '$1-$2', '$<n1>|$<n9>', "$`.$'", '$$1', '$10$01', '$<x'];

describe('regex matcher against the platform RegExp', () => {
	afterEach(() => tuneMemo());

	it('agrees on random patterns and inputs, unrolled and on counters', () => {
		const rng = random(20261120);
		let compared = 0;
		for (let round = 0; round < 3000; round++) {
			const source = pattern(rng, 0, { count: 0 });
			const ignoreCase = rng.next() < 0.3;
			let native: RegExp;
			try {
				native = new RegExp(source, ignoreCase ? 'ui' : 'u');
			} catch {
				expect(() => compileRegex(source, ignoreCase), source).toThrow();
				continue;
			}
			const program = compileRegex(source, ignoreCase);
			const counted = compile(parse(source, ignoreCase), false);
			const global = new RegExp(source, `${native.flags}g`);
			for (let k = 0; k < 4; k++) {
				const text = input(rng);
				const want = [...text.matchAll(global)];
				if (want.some((m) => insidePair(text, m.index))) continue;
				const label = `/${source}/${ignoreCase ? 'i' : ''} on ${JSON.stringify(text)}`;
				const exec = expected(native.exec(text));
				for (const memo of MEMO) {
					const after = memo ? memo(rng) : undefined;
					tuneMemo(
						after === undefined ? undefined : () => after,
						after === undefined ? undefined : 1,
					);
					resetCalcBudget();
					const started = performance.now();
					const first = actual(text, search(program, text, 0), program.groupCount);
					const all = [...matches(program, text)].map((m) => text.slice(m[0], m[1]));
					expect(performance.now() - started, label).toBeLessThan(250);
					expect(first, `${label}, memo after ${after}`).toEqual(exec);
					expect(all, `${label}, memo after ${after}`).toEqual(want.map((m) => m[0]));
				}
				expect(actual(text, search(counted, text, 0), counted.groupCount), label).toEqual(exec);
				const template = rng.pick(TEMPLATES);
				expect(REPLACE?.([text, source, template, 0, ignoreCase ? 1 : 0], {} as never), label).toBe(
					text.replace(global, template),
				);
				compared++;
			}
		}
		expect(compared).toBeGreaterThan(5000);
	});
});
