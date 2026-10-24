/**
 * @fileoverview What ONE theme part carries, read without touching the runtime's theme state, and
 * its colour and font schemes in the shape `PptxTheme` carries them.
 *
 * The load pipeline reads every part a slide master uses into a {@link ThemePartModel} and caches
 * it (`parseThemePart`); `getThemePartSchemes` reads any part the same way and only hands its
 * schemes out; switching a master to another part (`setSlideMasterTheme`) adopts its model.
 */
import type {
	PptxThemeColorScheme,
	PptxThemeFontGroup,
	PptxThemeFontScheme,
	PptxThemeFormatScheme,
	PptxThemeObjectDefaults,
} from '../../types';
import { THEME_COLOR_SCHEME_KEYS } from './theme-scheme-edit';

/** One theme part, parsed. */
export interface ThemePartModel {
	/** `a:theme @name`, `a:clrScheme @name` and `a:fontScheme @name`, when present. */
	name: string | undefined;
	colorSchemeName: string | undefined;
	fontSchemeName: string | undefined;
	/**
	 * The twelve scheme slots over the schema defaults. The `tx1`/`bg1`/`tx2`/`bg2` aliases are not
	 * baked in: they resolve through each master's own `p:clrMap` at lookup time.
	 */
	colorMap: Record<string, string>;
	/** `mj-lt`/`mj-ea`/`mj-cs` and `mn-lt`/`mn-ea`/`mn-cs`. */
	fontMap: Record<string, string>;
	/** Per-script faces (`<a:font script="..."/>`) of the major and minor fonts. */
	majorFontScripts: Record<string, string>;
	minorFontScripts: Record<string, string>;
	/** Raw subtrees the save re-emits verbatim; `undefined` when the part has none. */
	objectDefaults: PptxThemeObjectDefaults | undefined;
	extraClrSchemeLst: unknown;
	custClrLst: unknown;
	extLst: unknown;
	formatScheme: PptxThemeFormatScheme | undefined;
}

/** The twelve scheme slots of a colour map (an absent slot is an empty string). */
export function colorSchemeFromMap(
	colorMap: Readonly<Record<string, string>>,
): PptxThemeColorScheme {
	const scheme: Record<string, string> = {};
	for (const key of THEME_COLOR_SCHEME_KEYS) {
		scheme[key] = colorMap[key] || '';
	}
	return scheme as unknown as PptxThemeColorScheme;
}

function fontGroup(
	fontMap: Readonly<Record<string, string>>,
	prefix: 'mj' | 'mn',
	byScript: Readonly<Record<string, string>>,
): PptxThemeFontGroup {
	const latin = fontMap[`${prefix}-lt`];
	const eastAsia = fontMap[`${prefix}-ea`];
	const complexScript = fontMap[`${prefix}-cs`];
	return {
		...(latin ? { latin } : {}),
		...(eastAsia ? { eastAsia } : {}),
		...(complexScript ? { complexScript } : {}),
		...(Object.keys(byScript).length > 0 ? { byScript: { ...byScript } } : {}),
	};
}

/** The major and minor fonts of a font map, with their per-script faces. */
export function fontSchemeFromMap(
	fontMap: Readonly<Record<string, string>>,
	majorByScript: Readonly<Record<string, string>>,
	minorByScript: Readonly<Record<string, string>>,
): PptxThemeFontScheme {
	return {
		majorFont: fontGroup(fontMap, 'mj', majorByScript),
		minorFont: fontGroup(fontMap, 'mn', minorByScript),
	};
}
