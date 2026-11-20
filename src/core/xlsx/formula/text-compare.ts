// How Excel compares text, established against Excel (see __fixtures__/excel-p4-cases.txt). Two
// rules apply, depending on the function:
//
// - Linguistic (the comparison operators, criteria such as COUNTIF, XMATCH, XLOOKUP, UNIQUE, SORT,
//   SWITCH): case-insensitive and accent-, width- and kana-sensitive. "ß" equals "ss", "æ" "ae",
//   "œ" "oe", a ligature its letters ("ﬁ" is "fi") and final sigma sigma. Hyphens and apostrophes
//   only break ties ("coop" < "co-op"), symbols sort before digits and digits before letters.
// - Character by character (MATCH, VLOOKUP, HLOOKUP, wildcards, SEARCH): each character matches
//   its simple case folding, so "ẞ" is "ß" and "ς" is "σ", but "ß" is not "ss".

/** Which of the two rules a lookup compares text by. */
export type TextRule = 'chars' | 'linguistic';

/** The text two values must share to be equal under `rule`. */
export const ruleKey = (text: string, rule: TextRule): string =>
	rule === 'chars' ? foldText(text) : textKey(text);

/** Excel's order of the ASCII symbols, after the space (the Sort documentation lists it). */
const SYMBOL_ORDER = ' !"#$%&()*,./:;?@[\\]^_`{|}~+<=>';
/** Sorted only after every other difference: they never separate otherwise equal text. */
const IGNORABLE = "'-";

const SYMBOL = 0x100;
const DIGIT = 0x200;
const LETTER = 0x300;
const OTHER_SYMBOL = 0x10_0000;
const OTHER_LETTER = 0x20_0000;

/** Primary weight of each ASCII code unit (folded), `-1` for the ignorable ones. */
const ASCII_WEIGHT = new Int32Array(128);
for (let code = 0; code < 128; code++) {
	const ch = String.fromCharCode(code);
	const lower = ch.toLowerCase().charCodeAt(0);
	if (IGNORABLE.includes(ch)) ASCII_WEIGHT[code] = -1;
	else if (code >= 0x30 && code <= 0x39) ASCII_WEIGHT[code] = DIGIT + code;
	else if (lower >= 0x61 && lower <= 0x7a) ASCII_WEIGHT[code] = LETTER + lower;
	else {
		const at = SYMBOL_ORDER.indexOf(ch);
		ASCII_WEIGHT[code] = at >= 0 ? SYMBOL + at : code;
	}
}

/** Expansions a linguistic comparison makes, after case folding. */
const EXPANSIONS: ReadonlyMap<string, string> = new Map([
	['ß', 'ss'],
	['æ', 'ae'],
	['œ', 'oe'],
	['ﬀ', 'ff'],
	['ﬁ', 'fi'],
	['ﬂ', 'fl'],
	['ﬃ', 'ffi'],
	['ﬄ', 'ffl'],
	['ﬅ', 'st'],
	['ﬆ', 'st'],
]);

const PRINTABLE_ASCII = /^[\x20-\x7e]*$/;
const MARK = /\p{M}/u;
const SYMBOLIC = /[\p{P}\p{S}\p{Z}\p{C}]/u;

/** Folded code units outside ASCII, as met: at most one entry per code unit. */
const FOLDED = new Map<string, string>();

/**
 * The simple case folding of one UTF-16 code unit, which never changes the length: "ẞ" is "ß",
 * "ς" and "Σ" are "σ", "ß" and "ı" stay as they are.
 */
export function foldChar(ch: string): string {
	const code = ch.charCodeAt(0);
	if (code < 0x80) return code >= 0x41 && code <= 0x5a ? String.fromCharCode(code + 32) : ch;
	let folded = FOLDED.get(ch);
	if (folded === undefined) {
		const upper = ch.toUpperCase();
		const lower = (upper.length === ch.length ? upper : ch).toLowerCase();
		folded = ch === 'ı' || lower.length !== ch.length ? ch : lower;
		FOLDED.set(ch, folded);
	}
	return folded;
}

/** {@link foldChar} applied to every code unit: what character-by-character matching compares. */
export function foldText(text: string): string {
	if (PRINTABLE_ASCII.test(text)) return text.toLowerCase();
	let out = '';
	for (let i = 0; i < text.length; i++) out += foldChar(text[i] as string);
	return out;
}

/** Linguistic form: folded, decomposed (accents become marks) and expanded. */
function linguisticForm(text: string): string {
	let out = '';
	for (const ch of text.normalize('NFD')) {
		const folded = ch.length === 1 ? foldChar(ch) : ch;
		out += EXPANSIONS.get(folded) ?? folded;
	}
	return out;
}

/**
 * A key that two texts share exactly when the linguistic comparison finds them equal, for
 * hashing (criteria indexes, UNIQUE).
 */
export function textKey(text: string): string {
	return PRINTABLE_ASCII.test(text) ? text.toLowerCase() : linguisticForm(text);
}

/** The letter a full-width form or a katakana stands for, which sorts with it. */
function primaryChar(ch: string): { base: string; variant: number } {
	const code = ch.codePointAt(0) ?? 0;
	if (code >= 0xff01 && code <= 0xff5e) return { base: foldChar(ch.normalize('NFKC')), variant: 1 };
	if (code >= 0x30a1 && code <= 0x30f6)
		return { base: String.fromCharCode(code - 0x60), variant: -1 };
	return { base: ch, variant: 0 };
}

function weightOf(ch: string): number {
	const code = ch.codePointAt(0) ?? 0;
	if (code < 0x80) return ASCII_WEIGHT[code] ?? code;
	return (SYMBOLIC.test(ch) ? OTHER_SYMBOL : OTHER_LETTER) + code;
}

/** Each level of a linguistic comparison, most significant first. */
interface Levels {
	primary: number[];
	/** The accents after each primary character. */
	marks: string[];
	/** Hyphens and apostrophes with their positions among the primary characters. */
	ignorable: number[];
	/** Full-width and katakana forms by position. */
	variant: number[];
}

/** Recent texts' levels: a sort or a binary search compares the same texts many times. */
const LEVELS = new Map<string, Levels>();
const MAX_CACHED_LEVELS = 50_000;

function levels(text: string): Levels {
	let out = LEVELS.get(text);
	if (out) return out;
	if (LEVELS.size >= MAX_CACHED_LEVELS) LEVELS.clear();
	out = computeLevels(text);
	LEVELS.set(text, out);
	return out;
}

function computeLevels(text: string): Levels {
	const out: Levels = { primary: [], marks: [], ignorable: [], variant: [] };
	for (const ch of linguisticForm(text)) {
		if (MARK.test(ch)) {
			const last = out.marks.length - 1;
			if (last >= 0) out.marks[last] += ch;
			continue;
		}
		const { base, variant } = primaryChar(ch);
		const weight = weightOf(base);
		if (weight < 0) {
			out.ignorable.push(out.primary.length * 2 + (ch === '-' ? 1 : 0));
			continue;
		}
		out.primary.push(weight);
		out.marks.push('');
		out.variant.push(variant);
	}
	return out;
}

function compareLists<T>(a: readonly T[], b: readonly T[]): number {
	const n = Math.min(a.length, b.length);
	for (let i = 0; i < n; i++) {
		const x = a[i] as T;
		const y = b[i] as T;
		if (x !== y) return x < y ? -1 : 1;
	}
	return a.length === b.length ? 0 : a.length < b.length ? -1 : 1;
}

const isAscii = (code: number): boolean => code < 0x80;

/**
 * The primary level over the ASCII prefix of both texts, without allocating: the order when that
 * prefix decides it, 0 when both texts are ASCII and tie, undefined when a character outside ASCII
 * is reached first (accents, expansions and other scripts need the full comparison).
 */
function compareAsciiPrefix(a: string, b: string): number | undefined {
	let i = 0;
	let j = 0;
	for (;;) {
		while (
			i < a.length &&
			isAscii(a.charCodeAt(i)) &&
			(ASCII_WEIGHT[a.charCodeAt(i)] as number) < 0
		)
			i++;
		while (
			j < b.length &&
			isAscii(b.charCodeAt(j)) &&
			(ASCII_WEIGHT[b.charCodeAt(j)] as number) < 0
		)
			j++;
		if ((i < a.length && !isAscii(a.charCodeAt(i))) || (j < b.length && !isAscii(b.charCodeAt(j))))
			return undefined;
		if (i >= a.length || j >= b.length) break;
		const wa = ASCII_WEIGHT[a.charCodeAt(i)] as number;
		const wb = ASCII_WEIGHT[b.charCodeAt(j)] as number;
		if (wa !== wb) return wa < wb ? -1 : 1;
		i++;
		j++;
	}
	if (i < a.length) return 1;
	if (j < b.length) return -1;
	return 0;
}

/** Excel's case-insensitive linguistic order of two texts: negative, 0 or positive. */
export function compareText(a: string, b: string): number {
	if (a === b) return 0;
	const prefix = compareAsciiPrefix(a, b);
	if (prefix) return prefix;
	if (prefix === 0 && a.toLowerCase() === b.toLowerCase()) return 0;
	const la = levels(a);
	const lb = levels(b);
	return (
		compareLists(la.primary, lb.primary) ||
		compareLists(la.marks, lb.marks) ||
		compareLists(la.ignorable, lb.ignorable) ||
		compareLists(la.variant, lb.variant)
	);
}
