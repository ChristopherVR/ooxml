import { FONT_SUBSTITUTION_MAP } from './font-substitution-map';
import { stripFontStyleSuffix } from './font-style-name';
import { classifyPanose } from './panose';
export { FONT_SUBSTITUTION_MAP } from './font-substitution-map';
export { stripFontStyleSuffix } from './font-style-name';
export * from './panose';

const canonicalSubstitutes = new Map(
	Object.entries(FONT_SUBSTITUTION_MAP).map(([name, fonts]) => [name.toLowerCase(), fonts]),
);
function lookupFontSubstitutes(name: string): readonly string[] | undefined {
	return canonicalSubstitutes.get(name.trim().toLowerCase());
}
const GENERIC_FAMILY_CONCRETE_FONTS: Record<string, readonly string[]> = {
	serif: ['Times New Roman', 'Georgia', 'serif'],
	'sans-serif': ['Arial', 'Helvetica', 'sans-serif'],
	monospace: ['Courier New', 'Consolas', 'monospace'],
	cursive: ['Comic Sans MS', 'cursive'],
	fantasy: ['Impact', 'fantasy'],
};

export function getSubstituteFontFamily(fontName: string, panose?: number[]): string {
	const trimmed = fontName.trim();
	if (!trimmed) {
		return 'sans-serif';
	}
	return buildFontFamilyString(trimmed, getSubstituteFonts(trimmed, panose));
}

/**
 * Get only the fallback fonts for a given font name (without the
 * original font name itself). Useful when building font-family chains
 * where the primary font is already included.
 *
 * @param fontName The font family name as referenced in the PPTX.
 * @param panose Optional PANOSE classification array (10 bytes).
 * @returns Array of fallback font names (may include generic CSS families).
 */
export function getSubstituteFonts(fontName: string, panose?: number[]): readonly string[] {
	const trimmed = fontName.trim();
	if (!trimmed) {
		return ['sans-serif'];
	}

	// A full font name ("Bebas Neue Regular") also tries its family base
	// ("Bebas Neue"), both as an installed-font candidate and as the key
	// into the substitution map.
	const baseName = stripFontStyleSuffix(trimmed);
	const basePrefix: string[] = baseName ? [baseName] : [];

	// Strategy 1: Direct substitution map (full name first, then base)
	const directSubs =
		lookupFontSubstitutes(trimmed) ?? (baseName ? lookupFontSubstitutes(baseName) : undefined);
	if (directSubs) {
		return [...basePrefix, ...directSubs];
	}

	// Strategy 2: PANOSE-based classification
	if (panose && panose.length >= 4) {
		const genericFamily = classifyPanose(panose);
		const concreteFonts = GENERIC_FAMILY_CONCRETE_FONTS[genericFamily];
		if (concreteFonts) {
			return [...basePrefix, ...concreteFonts];
		}
	}

	// Strategy 3: Default fallback
	return [...basePrefix, 'sans-serif'];
}

/**
 * Check whether a font name has a known direct substitution.
 *
 * @param fontName The font family name.
 * @returns `true` if the font is in the direct substitution map.
 */
export function hasDirectSubstitution(fontName: string): boolean {
	return lookupFontSubstitutes(fontName) !== undefined;
}

/* ------------------------------------------------------------------ */
/*  CSS string builder                                                */
/* ------------------------------------------------------------------ */

/** CSS generic font family keywords that should NOT be quoted. */
const CSS_GENERIC_FAMILIES = new Set([
	'serif',
	'sans-serif',
	'monospace',
	'cursive',
	'fantasy',
	'system-ui',
	'ui-serif',
	'ui-sans-serif',
	'ui-monospace',
	'ui-rounded',
	'math',
	'emoji',
	'fangsong',
]);

/**
 * Build a CSS `font-family` value string from a primary font name
 * and an ordered list of fallback fonts.
 *
 * Generic CSS family keywords (serif, sans-serif, etc.) are not quoted.
 * All other font names are wrapped in double quotes.
 *
 * @param primary The primary font family name.
 * @param fallbacks Ordered array of fallback font names / generic families.
 * @returns A properly formatted CSS font-family string.
 */
export function buildFontFamilyString(primary: string, fallbacks: readonly string[]): string {
	const parts: string[] = [quoteFontName(primary)];
	for (const fb of fallbacks) {
		const quoted = quoteFontName(fb);
		// Avoid duplicates in the chain
		if (!parts.includes(quoted)) {
			parts.push(quoted);
		}
	}
	return parts.join(', ');
}

/**
 * Quote a font family name for CSS use, unless it is a generic
 * CSS family keyword.
 */
function quoteFontName(name: string): string {
	const trimmed = name.trim();
	if (CSS_GENERIC_FAMILIES.has(trimmed.toLowerCase())) {
		return trimmed;
	}
	return `"${trimmed.replace(/["\\\r\n\f]/g, '')}"`;
}
