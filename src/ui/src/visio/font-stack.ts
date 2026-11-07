import { getSubstituteFontFamily } from 'ooxml-core/text';

/** Distinct families remembered; a drawing uses a handful, so this only bounds hostile input. */
const MAX_CACHED = 512;
const stacks = new Map<string, string>();

/**
 * The CSS `font-family` list text is painted and measured with: the drawing's family first, then
 * its metric-compatible substitutes (Calibri, then Carlito; Cambria, then Caladea) and a generic
 * family. A reader without the Office font gets the clone, installed or loaded by the webfont
 * fallback (`googleFonts`), so line breaks follow the saved metrics instead of the browser's
 * default face. Painting and the canvas measurer must use the same list, or wrapping and the
 * drawn glyphs disagree.
 */
export function visioFontStack(family: string): string {
	const cached = stacks.get(family);
	if (cached !== undefined) return cached;
	const stack = getSubstituteFontFamily(family);
	if (stacks.size >= MAX_CACHED) stacks.clear();
	stacks.set(family, stack);
	return stack;
}

/** The first family of a CSS `font-family` list, unquoted (the drawing's own family). */
export function primaryFontFamily(value: string): string {
	const trimmed = value.trim();
	const quote = trimmed[0];
	if (quote === '"' || quote === "'") {
		const end = trimmed.indexOf(quote, 1);
		return end < 0 ? trimmed.slice(1) : trimmed.slice(1, end);
	}
	return (trimmed.split(',')[0] ?? '').trim();
}
