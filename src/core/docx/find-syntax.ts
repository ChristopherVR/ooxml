// Word's Find and Replace codes, shared by the pattern matcher and the replacement text. Find
// searches one string per document: a paragraph mark is `\r` (Word's `^p`), a manual line break
// `\v`, a page break `\f`, and an object Find cannot see `\u0000`, which no match ever includes.

export const FIND_PARAGRAPH_MARK = '\r';
export const FIND_LINE_BREAK = '\v';
export const FIND_PAGE_BREAK = '\f';
export const FIND_OBJECT = '\u0000';

/** A search pattern or replacement text Word would reject. */
export class FindPatternError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'FindPatternError';
	}
}

/** The characters a regular expression needs escaped to match them literally. */
const SYNTAX_CHARACTERS = '^$\\.*+?()[]{}|/';

/** The regular-expression source that matches one literal character. */
export function escapeFindChar(char: string): string {
	return SYNTAX_CHARACTERS.includes(char) ? `\\${char}` : char;
}

const FIND_CODE_SOURCE: Readonly<Record<string, string>> = {
	p: FIND_PARAGRAPH_MARK,
	t: '\t',
	l: FIND_LINE_BREAK,
	m: FIND_PAGE_BREAK,
	s: ' ',
	'~': '‑',
	'-': '­',
	'^': '\\^',
	'?': '[^\\u0000\\r]',
	'#': '\\p{Nd}',
	$: '\\p{L}',
	w: '[^\\S\\r]',
};

/** The regular-expression source a Find code matches; `code` is the character after `^`. */
export function findCodeSource(code: string): string {
	const source = FIND_CODE_SOURCE[code.toLowerCase()];
	if (source === undefined) throw new FindPatternError(`Unknown search code ^${code}`);
	return source;
}
