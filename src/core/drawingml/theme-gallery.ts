import { FEATURED_THEMES, OFFICE_2007_THEME } from './theme-gallery-data';
import { ADDITIONAL_THEMES } from './theme-gallery-legacy';
import { FEATURED_THEME_FORMATS } from './theme-formats-featured';
import { LEGACY_THEME_FORMATS } from './theme-formats-legacy';
import { readCapturedFormatScheme } from './theme-format-capture';

export * from './theme-presets';
export { THEME_EFFECT_PRESETS, sameThemeEffectStyles } from './theme-effect-presets';

const formats = { ...FEATURED_THEME_FORMATS, ...LEGACY_THEME_FORMATS };

/** Featured themes followed by alphabetized themes from earlier Word releases. */
export const DOCUMENT_THEME_CATALOGUE = [
	...FEATURED_THEMES,
	OFFICE_2007_THEME,
	...ADDITIONAL_THEMES,
].map((theme) => ({ ...theme, formatScheme: readCapturedFormatScheme(formats[theme.id]) }));
