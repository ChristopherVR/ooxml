import { readCapturedFormatScheme } from './theme-format-capture';
import { GALLERY_EFFECT_FORMATS } from './theme-effect-formats';
import { FEATURED_THEME_FORMATS } from './theme-formats-featured';
import { drawingThemeFormatSchemeXml, type ThemeFormatScheme } from './format-scheme';
import { THEME_EFFECT_NAMES, type ThemeEffectPresetName } from './theme-effect-names';

const formats = {
	Office: FEATURED_THEME_FORMATS.office,
	'Office 2007-2010': FEATURED_THEME_FORMATS['office-2007-2010'],
	...GALLERY_EFFECT_FORMATS,
} satisfies Record<ThemeEffectPresetName, string>;

export const THEME_EFFECT_PRESETS = THEME_EFFECT_NAMES.map((name) => ({
	name,
	formatScheme: readCapturedFormatScheme(formats[name]),
}));

export function sameThemeFormatScheme(a: ThemeFormatScheme, b: ThemeFormatScheme): boolean {
	if (a.name !== b.name) return false;
	const signature = (scheme: ThemeFormatScheme) =>
		drawingThemeFormatSchemeXml(scheme)
			.replace(/ xmlns:a="[^"]+"/gu, '')
			.replace(/>\s+</gu, '><');
	return a === b || signature(a) === signature(b);
}
