import { CharMatcher, type Flags, literalMatcher } from './charset';

/** Zero-width tests on the position. */
export const Assertion = {
	Start: 0,
	End: 1,
	LineStart: 2,
	LineEnd: 3,
	Boundary: 4,
	NotBoundary: 5,
} as const;
export type Assertion = (typeof Assertion)[keyof typeof Assertion];

/** The tree of a JavaScript pattern with the `u` flag. Non-capturing groups leave no node. */
export type Node =
	| { readonly kind: 'char'; readonly matcher: CharMatcher }
	| { readonly kind: 'seq'; readonly items: readonly Node[] }
	| { readonly kind: 'alt'; readonly items: readonly Node[] }
	| { readonly kind: 'group'; readonly index: number; readonly body: Node }
	| {
			readonly kind: 'repeat';
			readonly body: Node;
			readonly min: number;
			readonly max: number;
			readonly greedy: boolean;
			/** The capture groups inside the body, `[first, last + 1)`; each iteration clears them. */
			readonly groups: readonly [number, number];
	  }
	| { readonly kind: 'assert'; readonly what: Assertion; readonly word: CharMatcher | null }
	| {
			readonly kind: 'look';
			readonly behind: boolean;
			readonly negate: boolean;
			readonly body: Node;
	  }
	| { readonly kind: 'backref'; readonly groups: number[]; readonly ignoreCase: boolean };

export interface Parsed {
	readonly root: Node;
	readonly groupCount: number;
	/** Group name to group indexes (several when alternatives reuse a name). */
	readonly names: ReadonlyMap<string, readonly number[]>;
}

/** A pattern this parser does not accept; callers report it as an invalid pattern. */
export class RegexSyntaxError extends Error {}

const SYNTAX = new Set('^$\\.*+?()[]{}|/');
const ESCAPES: Record<string, number> = { f: 12, n: 10, r: 13, t: 9, v: 11 };
const QUANTIFIER = /\{(\d+)(?:(,)(\d*))?\}/y;
const HEX4 = /[0-9a-fA-F]{4}/y;
/** Counts past any input length behave as unbounded, so they are clamped. */
const COUNT_CAP = 2 ** 31 - 1;

const count = (digits: string): number => Math.min(Number(digits), COUNT_CAP);

/** Parses `source` as JavaScript RegExp syntax with the `u` flag (and `i` when `ignoreCase`). */
export function parse(source: string, ignoreCase: boolean): Parsed {
	let at = 0;
	let groupCount = 0;
	const names = new Map<string, number[]>();
	const namedRefs: { groups: number[]; name: string }[] = [];
	let highestRef = 0;
	const cache = new Map<string, CharMatcher>();
	const fail = (): never => {
		throw new RegexSyntaxError();
	};

	const matcher = (key: string, make: () => CharMatcher, flags: Flags): CharMatcher => {
		const id = `${flags.i ? 'i' : ''}${flags.s ? 's' : ''}:${key}`;
		let found = cache.get(id);
		if (!found) cache.set(id, (found = make()));
		return found;
	};
	const set = (src: string, flags: Flags): Node => ({
		kind: 'char',
		matcher: matcher(src, () => new CharMatcher(-1, src, flags), flags),
	});
	const literal = (cp: number, flags: Flags): Node => ({
		kind: 'char',
		matcher: matcher(`#${cp}`, () => literalMatcher(cp, flags), flags),
	});
	const word = (flags: Flags) => matcher('\\w', () => new CharMatcher(-1, '\\w', flags), flags);

	const hex = (digits: string): number =>
		/^[0-9a-fA-F]+$/.test(digits) ? parseInt(digits, 16) : fail();
	const hex4 = (): number => {
		HEX4.lastIndex = at;
		if (!HEX4.test(source)) fail();
		at += 4;
		return parseInt(source.slice(at - 4, at), 16);
	};
	/** After `\u`: `{hex}`, or four hex digits, joining an escaped surrogate pair. */
	const unicodeEscape = (): number => {
		if (source[at] === '{') {
			const close = source.indexOf('}', at);
			const cp = close < 0 ? fail() : hex(source.slice(at + 1, close));
			at = close + 1;
			return cp > 0x10ffff ? fail() : cp;
		}
		const lead = hex4();
		if (lead >= 0xd800 && lead < 0xdc00 && source.startsWith('\\u', at)) {
			const save = at;
			at += 2;
			HEX4.lastIndex = at;
			if (HEX4.test(source)) {
				const trail = hex4();
				if (trail >= 0xdc00 && trail <= 0xdfff)
					return (lead - 0xd800) * 0x400 + trail - 0xdc00 + 0x10000;
			}
			at = save;
		}
		return lead;
	};
	const groupName = (): string => {
		let name = '';
		while (source[at] !== '>') {
			if (at >= source.length) fail();
			if (source[at] === '\\') {
				at++;
				if (source[at++] !== 'u') fail();
				name += String.fromCodePoint(unicodeEscape());
			} else {
				const cp = source.codePointAt(at) ?? fail();
				name += String.fromCodePoint(cp);
				at += cp > 0xffff ? 2 : 1;
			}
		}
		at++;
		return name || fail();
	};

	/** An escape outside a class, after the backslash. */
	const escape = (flags: Flags): Node => {
		const ch = source[at++] ?? fail();
		if (ch === 'b' || ch === 'B')
			return {
				kind: 'assert',
				what: ch === 'b' ? Assertion.Boundary : Assertion.NotBoundary,
				word: word(flags),
			};
		if ('dDsSwW'.includes(ch)) return set(`\\${ch}`, flags);
		if (ch === 'p' || ch === 'P') {
			const close = source[at] === '{' ? source.indexOf('}', at) : -1;
			if (close < 0) fail();
			const src = `\\${ch}${source.slice(at, close + 1)}`;
			at = close + 1;
			return set(src, flags);
		}
		if (ch === 'k') {
			if (source[at++] !== '<') fail();
			const ref = { groups: [] as number[], name: groupName() };
			namedRefs.push(ref);
			return { kind: 'backref', groups: ref.groups, ignoreCase: flags.i };
		}
		if (ch >= '1' && ch <= '9') {
			let digits = ch;
			while (/[0-9]/.test(source[at] ?? '')) digits += source[at++];
			highestRef = Math.max(highestRef, Number(digits));
			return { kind: 'backref', groups: [Number(digits)], ignoreCase: flags.i };
		}
		if (ch === '0') return /[0-9]/.test(source[at] ?? '') ? fail() : literal(0, flags);
		const control = ESCAPES[ch];
		if (control !== undefined) return literal(control, flags);
		if (ch === 'c') {
			const letter = source[at++] ?? '';
			return /^[A-Za-z]$/.test(letter) ? literal(letter.charCodeAt(0) % 32, flags) : fail();
		}
		if (ch === 'x') {
			at += 2;
			return literal(hex(source.slice(at - 2, at).padEnd(2, 'g')), flags);
		}
		if (ch === 'u') return literal(unicodeEscape(), flags);
		return SYNTAX.has(ch) ? literal(ch.charCodeAt(0), flags) : fail();
	};

	/** A bracketed class, handed whole to the platform (it validated the pattern first). */
	const charClass = (flags: Flags): Node => {
		let end = at + 1;
		if (source[end] === '^') end++;
		while (end < source.length && source[end] !== ']') end += source[end] === '\\' ? 2 : 1;
		if (end >= source.length) fail();
		const src = source.slice(at, end + 1);
		at = end + 1;
		return set(src, flags);
	};

	const group = (flags: Flags): { node: Node; quantifiable: boolean } => {
		at++;
		if (source[at] !== '?') {
			const index = ++groupCount;
			return { node: { kind: 'group', index, body: closing(flags) }, quantifiable: true };
		}
		const next = source[at + 1];
		const behind = next === '<' && (source[at + 2] === '=' || source[at + 2] === '!');
		if (next === '=' || next === '!' || behind) {
			const negate = source[at + (behind ? 2 : 1)] === '!';
			at += behind ? 3 : 2;
			return { node: { kind: 'look', behind, negate, body: closing(flags) }, quantifiable: false };
		}
		if (next === '<') {
			at += 2;
			const name = groupName();
			const index = ++groupCount;
			names.set(name, [...(names.get(name) ?? []), index]);
			return { node: { kind: 'group', index, body: closing(flags) }, quantifiable: true };
		}
		at++;
		const on = /[ims]*/y;
		on.lastIndex = at;
		const add = on.exec(source)?.[0] ?? '';
		at += add.length;
		let remove = '';
		if (source[at] === '-') {
			on.lastIndex = ++at;
			remove = on.exec(source)?.[0] ?? '';
			at += remove.length;
		}
		if (source[at++] !== ':') fail();
		const pick = (flag: 'i' | 'm' | 's') =>
			add.includes(flag) ? true : remove.includes(flag) ? false : flags[flag];
		return { node: closing({ i: pick('i'), m: pick('m'), s: pick('s') }), quantifiable: true };
	};
	const closing = (flags: Flags): Node => {
		const body = disjunction(flags);
		if (source[at++] !== ')') fail();
		return body;
	};

	const term = (flags: Flags): { node: Node; quantifiable: boolean } => {
		const ch = source[at];
		if (ch === '^' || ch === '$') {
			at++;
			const what =
				ch === '^'
					? flags.m
						? Assertion.LineStart
						: Assertion.Start
					: flags.m
						? Assertion.LineEnd
						: Assertion.End;
			return { node: { kind: 'assert', what, word: null }, quantifiable: false };
		}
		if (ch === '\\') {
			at++;
			const node = escape(flags);
			return { node, quantifiable: node.kind !== 'assert' };
		}
		if (ch === '(') return group(flags);
		if (ch === '[') return { node: charClass(flags), quantifiable: true };
		if (ch === '.') {
			at++;
			return { node: set('.', flags), quantifiable: true };
		}
		if (ch === undefined || '*+?{}]'.includes(ch)) fail();
		const cp = source.codePointAt(at) ?? fail();
		at += cp > 0xffff ? 2 : 1;
		return { node: literal(cp, flags), quantifiable: true };
	};

	const quantifier = (): [number, number] | null => {
		const ch = source[at];
		if (ch === '*' || ch === '+' || ch === '?') {
			at++;
			return [ch === '+' ? 1 : 0, ch === '?' ? 1 : Infinity];
		}
		if (ch !== '{') return null;
		QUANTIFIER.lastIndex = at;
		const m = QUANTIFIER.exec(source) ?? fail();
		at += m[0].length;
		const min = count(m[1] ?? '');
		const max = !m[2] ? min : m[3] ? count(m[3]) : Infinity;
		if (max < min) fail();
		return [min, max >= COUNT_CAP ? Infinity : max];
	};

	const sequence = (flags: Flags): Node => {
		const items: Node[] = [];
		while (at < source.length && source[at] !== '|' && source[at] !== ')') {
			const before = groupCount;
			const { node, quantifiable } = term(flags);
			const q = quantifier();
			if (!q) {
				items.push(node);
				continue;
			}
			if (!quantifiable) fail();
			const greedy = source[at] !== '?';
			if (!greedy) at++;
			const groups = [before + 1, groupCount + 1] as const;
			items.push({ kind: 'repeat', body: node, min: q[0], max: q[1], greedy, groups });
		}
		return items.length === 1 ? (items[0] as Node) : { kind: 'seq', items };
	};
	const disjunction = (flags: Flags): Node => {
		const items = [sequence(flags)];
		while (source[at] === '|') {
			at++;
			items.push(sequence(flags));
		}
		return items.length === 1 ? (items[0] as Node) : { kind: 'alt', items };
	};

	const root = disjunction({ i: ignoreCase, m: false, s: false });
	if (at !== source.length || highestRef > groupCount) fail();
	for (const ref of namedRefs) ref.groups.push(...(names.get(ref.name) ?? fail()));
	return { root, groupCount, names };
}
