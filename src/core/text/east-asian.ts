/**
 * Character classes for East Asian line layout. `isEastAsianChar` came from the Word layout
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
