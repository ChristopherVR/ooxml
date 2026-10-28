/** Resolve installed family names from authored full font names. */
const FONT_STYLE_SUFFIX_TOKENS = new Set([
	'regular',
	'normal',
	'book',
	'roman',
	'thin',
	'hairline',
	'extralight',
	'ultralight',
	'demilight',
	'semilight',
	'light',
	'medium',
	'demibold',
	'semibold',
	'bold',
	'extrabold',
	'ultrabold',
	'heavy',
	'black',
	'extrablack',
	'ultra',
	'italic',
	'oblique',
]);

/**
 * Prefixes that, written as a separate word, belong to the weight token
 * after them: "Semi Bold", "Extra Light", "Ultra Black".
 */
const FONT_WEIGHT_PREFIX_TOKENS = new Set(['semi', 'demi', 'extra', 'ultra']);

/** Weight tokens a {@link FONT_WEIGHT_PREFIX_TOKENS} entry can precede. */
const FONT_PREFIXED_WEIGHT_TOKENS = new Set(['bold', 'light', 'black']);

/**
 * The family name with trailing weight/style tokens removed (up to two,
 * e.g. "Semibold Italic", plus the prefix of a spaced weight such as
 * "Semi Bold"), or `undefined` when nothing was stripped.
 *
 * The full name always stays FIRST in any font chain; the stripped base
 * is only an additional candidate, so genuine families that end in a
 * weight word ("Arial Black") still win when installed.
 */
export function stripFontStyleSuffix(fontName: string): string | undefined {
	let name = fontName.trim();
	let stripped = false;
	let previous = '';
	let tokens = 0;
	for (;;) {
		const idx = name.lastIndexOf(' ');
		if (idx <= 0) {
			break;
		}
		const token = name
			.slice(idx + 1)
			.toLowerCase()
			.replace(/-/g, '');
		const prefix =
			FONT_WEIGHT_PREFIX_TOKENS.has(token) && FONT_PREFIXED_WEIGHT_TOKENS.has(previous);
		if (!prefix && (tokens >= 2 || !FONT_STYLE_SUFFIX_TOKENS.has(token))) {
			break;
		}
		name = name.slice(0, idx).trim();
		stripped = true;
		previous = prefix ? '' : token;
		if (!prefix) {
			tokens++;
		}
	}
	return stripped && name.length > 0 ? name : undefined;
}
