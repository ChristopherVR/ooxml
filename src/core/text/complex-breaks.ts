/**
 * Line-break opportunities inside Thai, Lao, Khmer and Myanmar text, which is written without
 * spaces between words (UAX #14 line breaking class SA, "complex context dependent").
 *
 * Word 16 does not break these scripts at dictionary word boundaries: a long unspaced run fills
 * each line and breaks between two grapheme clusters (UAX #29), mid-word (measured with
 * `Range.Information(wdFirstCharacterLineNumber)` on Thai, Lao, Khmer and Myanmar paragraphs,
 * with and without a `w:lang`, on an install with no Thai or other South-East Asian editing
 * language enabled; Word with one installed may break at word boundaries). A break therefore never falls inside a cluster, so a base letter
 * keeps its vowel signs, tone marks and medials. Two kinds of boundary that UAX #29 allows are
 * kept together because breaking there would split a syllable that is drawn as one unit (Word's
 * output never broke at them):
 *
 *  - after a Thai or Lao leading vowel (written before the consonant it follows in speech);
 *  - after a Khmer COENG (U+17D2) or a Myanmar VIRAMA (U+1039), which stack the next consonant.
 */

/** True for a character of a script that UAX #14 places in class SA (Thai, Lao, Khmer, Myanmar). */
export function isComplexBreakChar(ch: string): boolean {
	const cp = ch.codePointAt(0) ?? 0;
	return (
		(cp >= 0x0e00 && cp <= 0x0eff) || // Thai, Lao
		(cp >= 0x1000 && cp <= 0x109f) || // Myanmar
		(cp >= 0x1780 && cp <= 0x17ff) || // Khmer
		(cp >= 0x19e0 && cp <= 0x19ff) || // Khmer Symbols
		(cp >= 0xa9e0 && cp <= 0xa9ff) || // Myanmar Extended-B
		(cp >= 0xaa60 && cp <= 0xaa7f) // Myanmar Extended-A
	);
}

/** Thai and Lao vowels written before their consonant, which never end a line. */
function isLeadingVowel(cp: number): boolean {
	return (cp >= 0x0e40 && cp <= 0x0e44) || (cp >= 0x0ec0 && cp <= 0x0ec4);
}

/** Signs that stack the following consonant below the current one. */
function isStacker(cp: number): boolean {
	return cp === 0x17d2 || cp === 0x1039;
}

/**
 * Whether a line may break between two adjacent grapheme clusters `before` and `after`, both of
 * complex-context scripts (see {@link isComplexBreakChar}).
 */
export function complexBreakAllowed(before: string, after: string): boolean {
	if (!before || !after) return false;
	if (!isComplexBreakChar(before) || !isComplexBreakChar(after)) return false;
	const last = before.codePointAt(before.length - 1) ?? 0;
	const lead = before.codePointAt(0) ?? 0;
	// A one-character cluster that is a leading vowel, or a cluster ending in a stacker.
	return !(before.length === 1 && isLeadingVowel(lead)) && !isStacker(last);
}
