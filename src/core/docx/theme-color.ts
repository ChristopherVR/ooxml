// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { applyThemePaletteVariant, toHex } from '../color/index';
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

/*
 * WordprocessingML transforms HSL luminance, as the shared Office theme palette does.
 * This differs from DrawingML's linear-light tint/shade transforms (MS-OI29500 17.3.2.6).
 */

/** Mixes toward black; `fraction` 0 = black, 1 = unchanged, per Word's `themeShade`. */
export function applyThemeShade(hex: string, fraction: number): string {
	if (!/^#?[0-9a-fA-F]{6}$/.test(hex)) return hex;
	return applyThemePaletteVariant(hex, {
		kind: 'darker',
		percent: (1 - fraction) * 100,
	}).toUpperCase();
}
/** Mixes toward white; `fraction` 0 = white, 1 = unchanged, per Word's `themeTint`. */
export function applyThemeTint(hex: string, fraction: number): string {
	if (!/^#?[0-9a-fA-F]{6}$/.test(hex)) return hex;
	return applyThemePaletteVariant(hex, {
		kind: 'lighter',
		percent: (1 - fraction) * 100,
	}).toUpperCase();
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
	if (ref.tint !== undefined) hex = applyThemeTint(hex, ref.tint);
	else if (ref.shade !== undefined) hex = applyThemeShade(hex, ref.shade);
	return hex;
}
