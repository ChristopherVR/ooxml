// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { linearToSrgb255, srgb255ToLinear, toHex } from '../color/index';
import { DEFAULT_THEME_COLOR_MAP } from '../drawingml/theme-model';
import type {
	ThemeCatalog,
	ThemeColorReference,
	ThemeColorSlot,
	ThemeColorToken,
} from './theme-model';
import { isStThemeColor, ST_ThemeColor } from './generated/wml-simple-types';

export const THEME_COLOR_TOKENS = ST_ThemeColor;
export function isThemeColorToken(value: string): value is ThemeColorToken {
	return isStThemeColor(value);
}

const TOKEN_TO_SLOT: Partial<Record<ThemeColorToken, ThemeColorSlot>> = {
	dark1: 'dk1',
	light1: 'lt1',
	dark2: 'dk2',
	light2: 'lt2',
	accent1: 'accent1',
	accent2: 'accent2',
	accent3: 'accent3',
	accent4: 'accent4',
	accent5: 'accent5',
	accent6: 'accent6',
	hyperlink: 'hlink',
	followedHyperlink: 'folHlink',
};

/** Resolves a `w:themeColor` token to the theme's `#RRGGBB` scheme color, before tint/shade. */
export function resolveThemeColorToken(
	token: ThemeColorToken,
	theme: ThemeCatalog,
): string | undefined {
	const logical: Partial<Record<ThemeColorToken, 'bg1' | 'tx1' | 'bg2' | 'tx2'>> = {
		background1: 'bg1',
		text1: 'tx1',
		background2: 'bg2',
		text2: 'tx2',
	};
	const logicalKey = logical[token];
	if (logicalKey) {
		// Word's conventional bg/tx defaults apply when settings.xml has no clrSchemeMapping.
		const slot = theme.colorMapping[logicalKey] ?? DEFAULT_THEME_COLOR_MAP[logicalKey];
		return theme.colors[slot];
	}
	const slot = TOKEN_TO_SLOT[token];
	return slot ? theme.colors[slot] : undefined;
}

function hexToRgb(hex: string): { r: number; g: number; b: number } | null {
	const normalized = hex.replace(/^#/, '');
	if (!/^[0-9a-fA-F]{6}$/.test(normalized)) return null;
	return {
		r: Number.parseInt(normalized.slice(0, 2), 16),
		g: Number.parseInt(normalized.slice(2, 4), 16),
		b: Number.parseInt(normalized.slice(4, 6), 16),
	};
}
/*
 * Linear-light shade/tint mixing (ECMA-376 Part 1, 20.1.2.3.30/20.1.2.3.32): PowerPoint and Word
 * mix a scheme color toward black/white in linear light rather than gamma-encoded sRGB, with the
 * IEC 61966-2-1 transfer function of the shared `color` area.
 */

/** Mixes toward black; `fraction` 0 = black, 1 = unchanged, per Word's `themeShade`. */
export function applyThemeShade(hex: string, fraction: number): string {
	const rgb = hexToRgb(hex);
	if (!rgb) return hex;
	const mix = (channel: number) => linearToSrgb255(srgb255ToLinear(channel) * fraction);
	return `#${toHex(mix(rgb.r))}${toHex(mix(rgb.g))}${toHex(mix(rgb.b))}`;
}
/** Mixes toward white; `fraction` 0 = white, 1 = unchanged, per Word's `themeTint`. */
export function applyThemeTint(hex: string, fraction: number): string {
	const rgb = hexToRgb(hex);
	if (!rgb) return hex;
	const mix = (channel: number) => linearToSrgb255(1 - (1 - srgb255ToLinear(channel)) * fraction);
	return `#${toHex(mix(rgb.r))}${toHex(mix(rgb.g))}${toHex(mix(rgb.b))}`;
}

/** Parses Word's two-hex-digit `themeTint`/`themeShade` byte into a [0,1] fraction. */
export function themeByteToFraction(value: string | undefined): number | undefined {
	if (value === undefined || !/^[0-9a-fA-F]{2}$/.test(value)) return undefined;
	return Number.parseInt(value, 16) / 255;
}
/** Serializes a [0,1] fraction back into Word's two-hex-digit byte form. */
export function fractionToThemeByte(fraction: number): string {
	return toHex(Math.round(Math.min(1, Math.max(0, fraction)) * 255));
}

/** Resolves a theme color reference (token + optional tint/shade) to a `#RRGGBB` hex string. */
export function resolveThemeColorReference(
	ref: ThemeColorReference,
	theme: ThemeCatalog,
): string | undefined {
	const base = resolveThemeColorToken(ref.token, theme);
	if (!base) return undefined;
	let hex = `#${base.replace(/^#/, '')}`;
	if (ref.shade !== undefined) hex = applyThemeShade(hex, ref.shade);
	if (ref.tint !== undefined) hex = applyThemeTint(hex, ref.tint);
	return hex;
}
