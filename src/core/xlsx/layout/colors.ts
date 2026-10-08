import type { Color, ThemePalette } from '../model';
import { applyTint, parseHex, toHexColor } from './tint';
import { themeColor } from '../../chart/render/theme-palette';

/**
 * The legacy 64-colour palette (`indexed` 0-63) as `RRGGBB`, followed by the system foreground
 * (64, window text) and background (65, window) colours. A workbook's `indexedColors` override can
 * be passed to {@link resolveColor} instead.
 */
export const INDEXED_COLORS: readonly string[] = [
	'000000',
	'FFFFFF',
	'FF0000',
	'00FF00',
	'0000FF',
	'FFFF00',
	'FF00FF',
	'00FFFF',
	'000000',
	'FFFFFF',
	'FF0000',
	'00FF00',
	'0000FF',
	'FFFF00',
	'FF00FF',
	'00FFFF',
	'800000',
	'008000',
	'000080',
	'808000',
	'800080',
	'008080',
	'C0C0C0',
	'808080',
	'9999FF',
	'993366',
	'FFFFCC',
	'CCFFFF',
	'660066',
	'FF8080',
	'0066CC',
	'CCCCFF',
	'000080',
	'FF00FF',
	'FFFF00',
	'00FFFF',
	'800080',
	'800000',
	'008080',
	'0000FF',
	'00CCFF',
	'CCFFFF',
	'CCFFCC',
	'FFFF99',
	'99CCFF',
	'FF99CC',
	'CC99FF',
	'FFCC99',
	'3366FF',
	'33CCCC',
	'99CC00',
	'FFCC00',
	'FF9900',
	'FF6600',
	'666699',
	'969696',
	'003366',
	'339966',
	'003300',
	'333300',
	'993300',
	'993366',
	'333399',
	'333333',
	'000000',
	'FFFFFF',
];

export { THEME_SLOTS, themeColor } from '../../chart/render/theme-palette';

/**
 * Resolves a SpreadsheetML colour to `#RRGGBB`. Explicit RGB wins, then theme (with Excel's tint),
 * then the indexed palette (64 and 65 are the system foreground and background). `auto` and a
 * missing colour give `fallback`. Alpha in `AARRGGBB` is ignored, as Excel does for cells.
 */
export function resolveColor(
	color: Color | undefined,
	theme: ThemePalette,
	fallback?: string,
	indexed: readonly string[] = INDEXED_COLORS,
): string | undefined {
	if (!color) return fallback;
	let base: string | undefined;
	if (color.rgb !== undefined) base = parseHex(color.rgb) ? color.rgb : undefined;
	else if (color.theme !== undefined) base = themeColor(theme, color.theme);
	else if (color.indexed !== undefined)
		base = indexed[color.indexed] ?? INDEXED_COLORS[color.indexed];
	if (base === undefined) return fallback;
	return applyTint(base, color.tint);
}

/** Linear RGB interpolation; Excel conditional scales truncate each channel delta toward zero. */
export function mixColors(
	a: string,
	b: string,
	t: number,
	rounding: 'nearest' | 'towardStart' = 'nearest',
): string {
	const ca = parseHex(a) ?? [0, 0, 0];
	const cb = parseHex(b) ?? [0, 0, 0];
	const k = Math.max(0, Math.min(1, t));
	return toHexColor(
		ca.map((v, i) => {
			const delta = ((cb[i] ?? 0) - v) * k;
			return v + (rounding === 'towardStart' ? Math.trunc(delta) : delta);
		}),
	);
}
