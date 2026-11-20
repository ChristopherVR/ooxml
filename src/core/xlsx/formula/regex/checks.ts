import {
	type CharMatcher,
	codePointAt,
	codePointBefore,
	isLineTerminator,
	literalMatcher,
	width,
} from './charset';
import { Assertion } from './parse';

/** Whether a zero-width assertion holds at `pos`; `word` is the `\w` matcher of `\b` and `\B`. */
export function assertion(
	input: string,
	kind: number,
	word: CharMatcher | undefined,
	pos: number,
): boolean {
	switch (kind) {
		case Assertion.Start:
			return pos === 0;
		case Assertion.End:
			return pos === input.length;
		case Assertion.LineStart:
			return pos === 0 || isLineTerminator(input.charCodeAt(pos - 1));
		case Assertion.LineEnd:
			return pos === input.length || isLineTerminator(input.charCodeAt(pos));
		default: {
			// No astral code point is a word character, so code units are enough.
			const before = pos > 0 && !!word?.test(input.charCodeAt(pos - 1));
			const after = pos < input.length && !!word?.test(input.charCodeAt(pos));
			return (before !== after) === (kind === Assertion.Boundary);
		}
	}
}

const folds = new Map<number, CharMatcher>();

/** Case-insensitive equality of two code points, as JavaScript's `i` flag folds them. */
function foldEqual(x: number, y: number): boolean {
	let m = folds.get(x);
	if (!m) folds.set(x, (m = literalMatcher(x, { i: true, m: false, s: false })));
	return m.test(y);
}

/**
 * The position after a backreference to the first set group of `groups` at `pos` (before it when
 * `back`), or -1. A group that did not take part matches empty.
 */
export function backref(
	input: string,
	regs: Int32Array,
	groups: readonly number[],
	back: boolean,
	ignoreCase: boolean,
	pos: number,
): number {
	const g = groups.find(
		(group) => (regs[2 * group] as number) >= 0 && (regs[2 * group + 1] as number) >= 0,
	);
	if (g === undefined) return pos;
	const s = regs[2 * g] as number;
	const e = regs[2 * g + 1] as number;
	if (!ignoreCase) {
		const from = back ? pos - (e - s) : pos;
		if (from < 0 || from + (e - s) > input.length) return -1;
		for (let i = 0; i < e - s; i++)
			if (input.charCodeAt(s + i) !== input.charCodeAt(from + i)) return -1;
		return back ? from : from + (e - s);
	}
	let i = back ? e : s;
	let j = pos;
	while (back ? i > s : i < e) {
		if (back ? j <= 0 : j >= input.length) return -1;
		const x = back ? codePointBefore(input, i) : codePointAt(input, i);
		const y = back ? codePointBefore(input, j) : codePointAt(input, j);
		if (x !== y && !foldEqual(x, y)) return -1;
		i += back ? -width(x) : width(x);
		j += back ? -width(y) : width(y);
	}
	return j;
}
