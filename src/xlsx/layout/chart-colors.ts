import { hslToRgb, rgbToHsl } from '../../color/index.js';
import type { ThemePalette } from '../model.js';
import { themeColor } from './colors.js';
import { parseHex, toHexColor } from './tint.js';

/**
 * Excel's automatic series colours: accent1..accent6, then the same accents darkened or
 * lightened for each further round of six (the `lumMod`/`lumOff` variations of the default
 * chart style: 60%, 80% + 20%, 80%, 60% + 40%, 50%, 70% + 30%, 70%, 50% + 50%).
 */
const VARIATIONS: readonly (readonly [lumMod: number, lumOff: number])[] = [
	[1, 0],
	[0.6, 0],
	[0.8, 0.2],
	[0.8, 0],
	[0.6, 0.4],
	[0.5, 0],
	[0.7, 0.3],
	[0.7, 0],
	[0.5, 0.5],
];

/** DrawingML `lumMod`/`lumOff` on HSL luminance. */
export function modulateLuminance(hex: string, lumMod: number, lumOff = 0): string {
	const rgb = parseHex(hex);
	if (!rgb) return '#000000';
	if (lumMod === 1 && lumOff === 0) return toHexColor(rgb);
	const hsl = rgbToHsl(...rgb);
	const l = Math.max(0, Math.min(1, hsl.l * lumMod + lumOff));
	const out = hslToRgb(hsl.h, hsl.s, l);
	return toHexColor([out.r, out.g, out.b]);
}

/** The automatic colour of series (or pie point) `index`, `#RRGGBB`. */
export function autoSeriesColor(theme: ThemePalette, index: number): string {
	const i = Math.max(0, Math.floor(index));
	const accent = themeColor(theme, 4 + (i % 6)) ?? '4472C4';
	const [lumMod, lumOff] = VARIATIONS[Math.floor(i / 6) % VARIATIONS.length] ?? [1, 0];
	return modulateLuminance(accent, lumMod, lumOff);
}
