/**
 * Character classes for East Asian line layout: where a line may break around ideographs, and
 * where Word adds space between East Asian and Western text (`w:autoSpaceDE`, `w:autoSpaceDN`,
 * ECMA-376 Part 1 17.3.1.2 and 17.3.1.3). `isEastAsianChar` came from the Word layout
 * engine's tokenizer (`docx/layout/text-breaks.ts`), which had taken its ranges from the
 * PowerPoint renderer's East Asian break helper (`src/ui/src/pptx/render/text-east-asian-breaks.ts`).
 */

/**
 * Whether `ch` is an East Asian (CJK, kana, Hangul, fullwidth) character, around which a line
 * may break (UAX #14 class ID and its relatives), subject to kinsoku (see `kinsoku.ts`).
 */
export function isEastAsianChar(ch: string): boolean {
	const cp = ch.codePointAt(0) ?? 0;
	return (
		(cp >= 0x1100 && cp <= 0x11ff) ||
		(cp >= 0x2e80 && cp <= 0x9fff) ||
		(cp >= 0xa960 && cp <= 0xa97f) ||
		(cp >= 0xac00 && cp <= 0xd7ff) ||
		(cp >= 0xf900 && cp <= 0xfaff) ||
		(cp >= 0xfe30 && cp <= 0xfe4f) ||
		(cp >= 0xff00 && cp <= 0xffef) ||
		(cp >= 0x20000 && cp <= 0x3ffff)
	);
}

const EAST_ASIAN_LETTER =
	/^[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
const LETTER = /^\p{L}/u;
const DIGIT = /^\p{Nd}/u;

/** An ideograph, kana or Hangul syllable: the East Asian side of automatic spacing. */
export function isEastAsianLetter(grapheme: string): boolean {
	return EAST_ASIAN_LETTER.test(grapheme);
}

/** Which automatic spacing applies at a boundary: Western letters (DE) or digits (DN). */
export type AutoSpaceKind = 'alpha' | 'digit';

function westernKind(grapheme: string): AutoSpaceKind | undefined {
	if (!grapheme || isEastAsianChar(grapheme)) return undefined;
	if (DIGIT.test(grapheme)) return 'digit';
	return LETTER.test(grapheme) ? 'alpha' : undefined;
}

/**
 * The automatic spacing Word inserts between two adjacent graphemes: `alpha` between an East
 * Asian letter and a Western letter (any script: Latin, Greek, Cyrillic), `digit` between an East
 * Asian letter and a digit, in either order. Punctuation on either side gets none (measured:
 * `日.日`, `日(a)日` and `日 a` are not spaced), and fullwidth letters count as East Asian.
 */
export function eastAsianAutoSpace(before: string, after: string): AutoSpaceKind | undefined {
	if (isEastAsianLetter(before)) return westernKind(after);
	if (isEastAsianLetter(after)) return westernKind(before);
	return undefined;
}

/**
 * The automatic space as a share of the preceding character's font size. Word 16 adds exactly a
 * quarter em after an East Asian letter (measured at 10, 12 and 20 pt, also when the two sides
 * differ in size). After a Western character the measured gap follows that character's size but
 * varies with its font, from 0.2 em (Times New Roman) to 0.3 em (Courier New); a quarter em is
 * used for both.
 */
export const AUTO_SPACE_EM = 0.25;
