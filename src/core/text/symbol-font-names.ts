/**
 * Which fonts are symbol-encoded, by name only: no glyph tables, so the document model and
 * layout can ask without loading `symbol-fonts.ts`. DOM-free.
 */

/** Fonts whose characters select glyphs by code rather than by Unicode meaning. */
export const OFFICE_SYMBOL_FONTS = [
	'Symbol',
	'Wingdings',
	'Wingdings 2',
	'Wingdings 3',
	'Webdings',
] as const;

const SYMBOL_ENCODED = new Set<string>(OFFICE_SYMBOL_FONTS.map((font) => font.toLowerCase()));

/** The first family of a CSS-like list, unquoted and in lower case. */
export function normaliseFamily(family: string | undefined): string {
	return (family ?? '')
		.split(',')[0]!
		.trim()
		.replace(/^["']|["']$/g, '')
		.toLowerCase();
}

/** Whether `family` is one of Office's symbol-encoded fonts (Symbol, Wingdings 1-3, Webdings). */
export function isSymbolEncodedFont(family: string | undefined): boolean {
	return SYMBOL_ENCODED.has(normaliseFamily(family));
}

/**
 * Whether `family` is probably symbol-encoded: one of Office's symbol fonts, or a font named like
 * the other common ones (Marlett, MT Extra, Monotype Sorts, MS Outlook, Bookshelf Symbol...).
 * Without the font file this is the best guess of how a code below U+0100 is meant.
 */
export function isLikelySymbolFont(family: string | undefined): boolean {
	const name = normaliseFamily(family);
	return (
		SYMBOL_ENCODED.has(name) ||
		// Not Segoe UI Symbol or Noto Sans Symbols, which are Unicode fonts.
		/dings|^symbol|bookshelf symbol|sorts|^marlett$|^mt extra$|^ms outlook$|^ms reference specialty$/.test(
			name,
		)
	);
}
