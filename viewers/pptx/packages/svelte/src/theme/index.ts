/**
 * Theme surface for the Svelte binding: a direct re-export of the shared,
 * framework-agnostic `ViewerTheme` system (types, defaults, CSS-variable
 * helpers, and the Vermilion presets), mirroring the Vue package.
 */
export type { ThemeCatalogEntry, ViewerTheme, ViewerThemeColors } from 'ooxml-ui/pptx';
export {
	defaultCssVars,
	defaultRadius,
	defaultThemeColors,
	resolveThemeCatalogEntry,
	THEME_CATALOG,
	themeToCssVars,
	vermilionDarkColors,
	vermilionDarkTheme,
	vermilionLightColors,
	vermilionLightTheme,
	vermilionRadius,
} from 'ooxml-ui/pptx';
