// The theme a new Word document is saved with, and the one theme fonts resolve against in a
// document without a theme part: Office's default colours and the shared default theme fonts
// (`OFFICE_THEME_FONTS`, drawingml), the one place to change them.
import { OFFICE_COLORS } from '../drawingml/office-theme-colors';
import { themeFontPair, type ThemeFontPair } from '../drawingml/office-theme-fonts';
import type { ThemeCatalog, ThemeColorSlot, ThemeFontRole } from './theme-model';

const SLOTS = [
	'lt1',
	'dk1',
	'lt2',
	'dk2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const satisfies readonly ThemeColorSlot[];

/** The default Office theme as a Word theme catalog, with `fonts` overriding its typefaces. */
export function defaultWordTheme(fonts?: Partial<ThemeFontPair>): ThemeCatalog {
	const { major, minor } = themeFontPair(fonts);
	return {
		name: 'Office Theme',
		colorSchemeName: 'Office',
		fontSchemeName: 'Office',
		colors: Object.fromEntries(SLOTS.map((slot, index) => [slot, OFFICE_COLORS[index]])),
		colorMapping: {},
		fonts: { major: { latin: major }, minor: { latin: minor } },
	};
}

/** The Latin face of theme font `role`: the document theme's, else the default theme's. */
export function themeLatinFont(
	theme: ThemeCatalog | undefined,
	role: ThemeFontRole,
): string | undefined {
	return theme ? theme.fonts[role]?.latin : themeFontPair()[role];
}
