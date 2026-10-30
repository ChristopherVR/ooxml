// Find (Word's Find and Replace dialog) as a matcher over the document's text. The text holds a
// paragraph mark as `\r` and every break or object as the codes in find-syntax.ts; this module
// only matches strings, so the UI maps each match back to document positions.
import { escapeFindChar, findCodeSource, FindPatternError } from './find-syntax';
import { wildcardSource } from './find-wildcards';

/** The Find options of the dialog's More >> section. */
export interface FindOptions {
	readonly matchCase?: boolean;
	readonly wholeWords?: boolean;
	readonly wildcards?: boolean;
	readonly matchPrefix?: boolean;
	readonly matchSuffix?: boolean;
	readonly ignorePunctuation?: boolean;
	readonly ignoreWhitespace?: boolean;
}

export interface FindMatch {
	readonly start: number;
	readonly end: number;
	readonly text: string;
	/** `groups[n]` is what the pattern's group n matched (`\n` in the replacement); `groups[0]` is the match. */
	readonly groups: readonly string[];
}

export interface FindPattern {
	readonly expression: RegExp;
	readonly groupCount: number;
}

const WORD_CHARACTERS = '\\p{L}\\p{M}\\p{N}_';
/** Text white space a literal search may skip, but not a paragraph mark, line or page break. */
const WHITE_SPACE_SKIP = '[^\\S\\r\\v\\f]';
const PUNCTUATION_SKIP = '\\p{P}';
/** A soft hyphen shows only where a line breaks, so a search skips over one inside a match. */
const SOFT_HYPHEN_SKIP = '\\u00AD';
const SOFT_HYPHEN = '­';

interface Atom {
	readonly source: string;
	/** The character a plain (not coded) atom searches for. */
	readonly literal?: string;
}

/** The atoms of a plain search: literal characters and Find codes. */
function literalAtoms(query: string): Atom[] {
	const atoms: Atom[] = [];
	for (let index = 0; index < query.length;) {
		const char = String.fromCodePoint(query.codePointAt(index) ?? 0);
		index += char.length;
		if (char === '^') {
			const code = String.fromCodePoint(query.codePointAt(index) ?? 0);
			if (!query.charAt(index)) throw new FindPatternError('The search code ^ ends the pattern');
			index += code.length;
			atoms.push({ source: findCodeSource(code) });
		} else if (char === SOFT_HYPHEN) {
			// A typed soft hyphen is not searched for; Find for an optional hyphen uses ^-.
			continue;
		} else {
			atoms.push({ source: escapeFindChar(char), literal: char });
		}
	}
	return atoms;
}

/** Whether a plain search leaves this atom out because the options ignore its character. */
function ignored(atom: Atom, options: FindOptions): boolean {
	const char = atom.literal;
	if (char === undefined) return false;
	if (options.ignoreWhitespace && /^[^\S\r\v\f]$/u.test(char)) return true;
	return Boolean(options.ignorePunctuation) && /^\p{P}$/u.test(char);
}

/**
 * The expression for a query and its options, or `null` when the query can match nothing. A
 * wildcard query ignores the options Word disables with wildcards (whole words, prefix, suffix and
 * the ignore options).
 */
export function compileFindPattern(query: string, options: FindOptions = {}): FindPattern | null {
	if (!query) return null;
	const flags = options.matchCase ? 'gu' : 'giu';
	if (options.wildcards) {
		const { source, groupCount } = wildcardSource(query);
		return { expression: new RegExp(source, flags), groupCount };
	}
	const atoms = literalAtoms(query).filter((atom) => !ignored(atom, options));
	if (!atoms.length) return null;
	const skips = [SOFT_HYPHEN_SKIP];
	if (options.ignoreWhitespace) skips.push(WHITE_SPACE_SKIP);
	if (options.ignorePunctuation) skips.push(PUNCTUATION_SKIP);
	const between = `(?:${skips.join('|')})*`;
	const body = atoms.map((atom) => atom.source).join(between);
	const before = options.wholeWords || options.matchPrefix ? `(?<![${WORD_CHARACTERS}])` : '';
	const after = options.wholeWords || options.matchSuffix ? `(?![${WORD_CHARACTERS}])` : '';
	return { expression: new RegExp(before + body + after, flags), groupCount: 0 };
}

/** The string offsets where a grapheme cluster starts or ends, so a match never splits one. */
function graphemeBoundaries(text: string): Set<number> {
	const boundaries = new Set<number>([0, text.length]);
	if (typeof Intl.Segmenter === 'function') {
		const segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
		for (const part of segmenter.segment(text)) boundaries.add(part.index);
		return boundaries;
	}
	for (let index = 0; index < text.length;) {
		const point = text.codePointAt(index) ?? 0;
		index += point > 0xffff ? 2 : 1;
		while (index < text.length && /\p{Mark}/u.test(text.charAt(index))) index++;
		boundaries.add(index);
	}
	return boundaries;
}

/**
 * Every non-overlapping match of the query in the text, in order. Empty matches are dropped, and
 * a match that starts or ends inside a grapheme cluster is not a match.
 */
export function findMatches(text: string, query: string, options: FindOptions = {}): FindMatch[] {
	const pattern = compileFindPattern(query, options);
	if (!pattern) return [];
	const boundaries = graphemeBoundaries(text);
	const matches: FindMatch[] = [];
	for (const found of text.matchAll(pattern.expression)) {
		const start = found.index;
		const end = start + found[0].length;
		if (end === start || !boundaries.has(start) || !boundaries.has(end)) continue;
		matches.push({
			start,
			end,
			text: found[0],
			groups: Array.from(found, (group) => group ?? ''),
		});
	}
	return matches;
}
