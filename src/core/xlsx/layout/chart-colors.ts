import { hslToRgb, rgbToHsl } from '../../color/index.js';
import type { ThemePalette } from '../model.js';
import { themeColor } from './colors.js';
import { parseHex, toHexColor } from './tint.js';
import {
	chartPaletteSeriesColor,
	findChartColorPalette,
	type ChartColorScheme,
} from '../../chart/color-palettes';
import { THEME_SLOTS } from './colors';

/** Workbook theme in DrawingML slot order, including missing-slot fallbacks. */
export function chartColorScheme(theme: ThemePalette): ChartColorScheme {
	return Object.fromEntries(
		THEME_SLOTS.map((slot, index) => [slot, `#${themeColor(theme, index) ?? '000000'}`]),
	) as unknown as ChartColorScheme;
}

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
	return chartPaletteSeriesColor(findChartColorPalette(10)!, i, i + 1, chartColorScheme(theme));
}
