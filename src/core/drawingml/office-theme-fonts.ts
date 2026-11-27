/**
 * The typefaces of the default Office theme ("Office", the font scheme Office 2013 to 2022 give a
 * new document): the one place every product's new-file path takes its theme fonts from. A new
 * document names them only in its theme part (`a:majorFont` / `a:minorFont`, or Visio's face
 * table) and refers to them through the theme elsewhere (`+mj-lt` / `+mn-lt`, Word's
 * `majorHAnsi` / `minorHAnsi`, Excel's `scheme="minor"`), as Office's own new files do. Nothing
 * here ships a font: a reader without these faces substitutes its own. A caller overrides them
 * per document through each writer's theme option; see `docs/reproducible-saves.md`.
 */

/** A theme's heading (`major`) and body (`minor`) Latin typefaces. */
export interface ThemeFontPair {
	major: string;
	minor: string;
}

/** The default theme fonts of a new document. */
export const OFFICE_THEME_FONTS: Readonly<ThemeFontPair> = Object.freeze({
	major: 'Calibri Light',
	minor: 'Calibri',
});

/** {@link OFFICE_THEME_FONTS} with a caller's overrides (an empty name keeps the default). */
export function themeFontPair(override?: {
	major?: string | undefined;
	minor?: string | undefined;
}): ThemeFontPair {
	return {
		major: override?.major || OFFICE_THEME_FONTS.major,
		minor: override?.minor || OFFICE_THEME_FONTS.minor,
	};
}
