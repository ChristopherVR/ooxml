/**
 * Colour contrast: WCAG 2.x relative luminance and contrast ratio, and readable-text selection.
 *
 * Renderers use `contrastTextColor` for fills whose text colour was left implicit: PowerPoint
 * stores no colour for a great many runs and resolves one at paint time from what is behind them.
 * Interface code uses `contrastRatio` and `readableTextColor`, which follow WCAG exactly, so a
 * theme can be checked against the 4.5:1 text and 3:1 non-text minimums.
 *
 * @module color/contrast
 */

/** WCAG 2.x 1.4.3 minimum for normal text (AA). */
export const WCAG_TEXT_CONTRAST = 4.5;
/** WCAG 2.x 1.4.11 minimum for control boundaries, focus indicators and other non-text parts (AA). */
export const WCAG_NON_TEXT_CONTRAST = 3;

/**
 * Parse a hex colour string into `[r, g, b]` in 0..255.
 * Accepts `#RGB`, `#RRGGBB`, `RGB` and `RRGGBB`; returns `null` for anything else.
 */
export function parseHexColor(hex: string): [number, number, number] | null {
	const cleaned = hex.trim().replace(/^#/u, '');
	if (!/^[\da-fA-F]+$/u.test(cleaned)) return null;
	const digits =
		cleaned.length === 3
			? cleaned
					.split('')
					.map((c) => c + c)
					.join('')
			: cleaned;
	if (digits.length !== 6) return null;
	return [
		Number.parseInt(digits.slice(0, 2), 16),
		Number.parseInt(digits.slice(2, 4), 16),
		Number.parseInt(digits.slice(4, 6), 16),
	];
}

/** Parse `#rgb`/`#rrggbb` into `[r, g, b]` (0..255); falls back to mid-grey. */
export function parseHex(hex: string): [number, number, number] {
	return parseHexColor(hex) ?? [128, 128, 128];
}

/**
 * WCAG 2.x relative luminance of an sRGB colour, 0 (black) to 1 (white).
 * @see https://www.w3.org/TR/WCAG21/#dfn-relative-luminance
 */
export function relativeLuminance(r: number, g: number, b: number): number {
	const linear = (c: number): number => {
		const s = c / 255;
		return s <= 0.04045 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4;
	};
	return 0.2126 * linear(r) + 0.7152 * linear(g) + 0.0722 * linear(b);
}

/**
 * WCAG 2.x contrast ratio between two hex colours, from 1 to 21. The order does not matter.
 * Returns `undefined` when either colour is not a hex colour.
 * @see https://www.w3.org/TR/WCAG21/#dfn-contrast-ratio
 */
export function contrastRatio(a: string, b: string): number | undefined {
	const rgbA = parseHexColor(a);
	const rgbB = parseHexColor(b);
	if (!rgbA || !rgbB) return undefined;
	const la = relativeLuminance(...rgbA);
	const lb = relativeLuminance(...rgbB);
	return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/**
 * The candidate with the highest WCAG contrast against `fill`. With the default white and black
 * candidates the result reaches at least 4.58:1 on any opaque fill. Returns the first candidate
 * when `fill` is not a hex colour.
 */
export function readableTextColor(
	fill: string,
	candidates: readonly [string, ...string[]] = ['#ffffff', '#000000'],
): string {
	let best = candidates[0];
	let bestRatio = -1;
	for (const candidate of candidates) {
		const ratio = contrastRatio(candidate, fill);
		if (ratio !== undefined && ratio > bestRatio) {
			best = candidate;
			bestRatio = ratio;
		}
	}
	return best;
}

/**
 * Pick a readable text colour (near-black or near-white) for a given fill, using a perceived-
 * brightness threshold. Kept for renderers that match PowerPoint's own choice; interface code
 * uses `readableTextColor`.
 */
export function contrastTextColor(fill: string): string {
	const [r, g, b] = parseHex(fill);
	const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
	return lum > 0.6 ? '#1a1a1a' : '#ffffff';
}
