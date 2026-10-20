/**
 * font-weight-names.ts: the weight a typeface name implies, for families that
 * are weight variants of a Google Fonts family.
 *
 * Decks often name a weight as its own family: "DM Sans Light", "Open Sans
 * SemiBold", "Montserrat ExtraBold". That is how static fonts register on
 * Windows, and PowerPoint stores the name as such. The Google Fonts CSS2 API
 * serves one family ("DM Sans") with a weight axis instead, so the face is
 * only reachable as `font-family: "DM Sans"; font-weight: 300`. Without this
 * mapping the catalogue lookup misses the name, nothing is requested, and
 * every such run renders at the regular weight of whatever the fallback chain
 * finds.
 *
 * The weight is applied ONLY when the base name is a catalogue family and the
 * full name is not. A full name the catalogue serves on its own ("Archivo
 * Black") is a separate single-weight family, and a name outside the
 * catalogue ("Arial Rounded MT Bold", a brand font) gives no reliable signal:
 * asking for weight 700 of a family whose only face is 400 makes the browser
 * synthesise a second, heavier bold. Those names keep the plain 400/700
 * mapping they always had.
 *
 * @module render/font-weight-names
 */
import { findGoogleFontsFamily } from './google-fonts-lookup';

/** Single-token weight names (hyphens removed, lower case). */
const WEIGHT_TOKENS: ReadonlyMap<string, number> = new Map([
	['thin', 100],
	['hairline', 100],
	['extralight', 200],
	['ultralight', 200],
	['light', 300],
	['regular', 400],
	['normal', 400],
	['book', 400],
	['medium', 500],
	['semibold', 600],
	['demibold', 600],
	['bold', 700],
	['extrabold', 800],
	['ultrabold', 800],
	['black', 900],
	['heavy', 900],
	['extrablack', 900],
	['ultrablack', 900],
]);

/** Prefixes that join the next token when the name spells them apart ("Semi Bold"). */
const WEIGHT_PREFIXES = new Set(['semi', 'demi', 'extra', 'ultra']);

/** Style tokens that may follow the weight ("Light Italic"); ignored here. */
const STYLE_TOKENS = new Set(['italic', 'oblique']);

/** A typeface name split into its family and the weight it names. */
export interface FontWeightName {
	/** The name without its trailing weight/style tokens. */
	family: string;
	/** CSS weight the trailing tokens name (400 when only a style is named). */
	weight: number;
}

/**
 * Split trailing weight (and style) tokens off a typeface name, or `null`
 * when the name ends in neither. Handles both run-together and spaced or
 * hyphenated spellings ("SemiBold", "Semi Bold", "Semi-Bold").
 */
export function parseFontWeightName(name: string): FontWeightName | null {
	const tokens = name.trim().split(/\s+/u);
	let weight: number | undefined;
	let consumed = 0;
	let end = tokens.length;
	// Optional trailing style token, then at most one weight (possibly split
	// across two tokens).
	if (end > 1 && STYLE_TOKENS.has(normaliseToken(tokens[end - 1]))) {
		end -= 1;
		consumed += 1;
	}
	if (end > 1) {
		const last = normaliseToken(tokens[end - 1]);
		const prefix = end > 2 ? normaliseToken(tokens[end - 2]) : '';
		const joined = WEIGHT_PREFIXES.has(prefix) ? WEIGHT_TOKENS.get(prefix + last) : undefined;
		if (joined !== undefined) {
			weight = joined;
			end -= 2;
			consumed += 2;
		} else if (WEIGHT_TOKENS.has(last)) {
			weight = WEIGHT_TOKENS.get(last);
			end -= 1;
			consumed += 1;
		}
	}
	if (consumed === 0 || end < 1) {
		return null;
	}
	return { family: tokens.slice(0, end).join(' '), weight: weight ?? 400 };
}

function normaliseToken(token: string): string {
	return token.toLowerCase().replace(/-/gu, '');
}

/** The Google Fonts family serving a typeface, and the weight it selects. */
export interface GoogleFontsFace {
	/** Canonical catalogue spelling of the family to request. */
	family: string;
	/**
	 * Weight the typeface name selects within `family`, or `undefined` when the
	 * name is the catalogue family itself.
	 */
	weight?: number;
}

/**
 * Which Google Fonts family serves `typeface`, if any: the family itself when
 * the catalogue lists the full name, otherwise its base family plus the
 * weight the trailing tokens name ("DM Sans Light" is "DM Sans" at 300).
 */
export function resolveGoogleFontsFace(typeface: string): GoogleFontsFace | null {
	const exact = findGoogleFontsFamily(typeface);
	if (exact !== null) {
		return { family: exact };
	}
	const parsed = parseFontWeightName(typeface);
	if (!parsed) {
		return null;
	}
	const base = findGoogleFontsFamily(parsed.family);
	return base === null ? null : { family: base, weight: parsed.weight };
}

/** Per-name cache: this runs once per rendered run. */
const impliedWeightCache = new Map<string, number | null>();

/**
 * The weight a typeface name implies for its regular face, or `null` when the
 * name implies none we can rely on (see the module comment).
 */
function impliedWeight(typeface: string): number | null {
	let cached = impliedWeightCache.get(typeface);
	if (cached === undefined) {
		cached = resolveGoogleFontsFace(typeface)?.weight ?? null;
		impliedWeightCache.set(typeface, cached);
	}
	return cached;
}

/**
 * The first family of a CSS `font-family` list ('"DM Sans Light", "DM Sans",
 * sans-serif'), unquoted, or the input itself when it is a bare name.
 */
function primaryFamily(familyOrList: string): string {
	const first = familyOrList.split(',')[0]?.trim() ?? '';
	const quoted = /^(["'])(.*)\1$/u.exec(first);
	return quoted ? quoted[2].replace(/\\(.)/gu, '$1') : first;
}

/**
 * CSS `font-weight` for text in `fontFamily` (a typeface name or a resolved
 * CSS `font-family` list, whose first entry is used).
 *
 * Bold text is 700. Regular text is 400, except for a weight variant of a
 * Google Fonts family, which renders at the weight its name implies ("DM Sans
 * Light" is 300). Bold text in such a family is at least 700, so "Black" stays
 * 900. Every place that renders or measures text calls this, so the canvas
 * measurement and the painted text agree.
 */
export function resolveFontWeight(
	fontFamily: string | undefined,
	bold: boolean | undefined,
): number {
	const implied = fontFamily ? impliedWeight(primaryFamily(fontFamily)) : null;
	const regular = implied ?? 400;
	return bold ? Math.max(700, regular) : regular;
}

/**
 * The weight as a canvas/CSS font-shorthand fragment: the `bold` keyword (or
 * none) for the usual 700/400, a number otherwise. Kept as keywords so the
 * common case serialises exactly as before.
 */
export function fontWeightShorthand(
	fontFamily: string | undefined,
	bold: boolean | undefined,
): string {
	const value = fontWeightAttribute(fontFamily, bold);
	return value === undefined ? '' : `${value} `;
}

/**
 * The weight as an SVG/CSS attribute value: omitted (`undefined`) for 400,
 * `bold` for 700, the number otherwise.
 */
export function fontWeightAttribute(
	fontFamily: string | undefined,
	bold: boolean | undefined,
): string | undefined {
	const weight = resolveFontWeight(fontFamily, bold);
	if (weight === 400) {
		return undefined;
	}
	return weight === 700 ? 'bold' : String(weight);
}
