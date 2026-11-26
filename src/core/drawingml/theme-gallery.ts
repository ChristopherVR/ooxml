import { themeMatrix } from './theme-matrix';
import { FEATURED_THEMES } from './theme-gallery-data';
import { ADDITIONAL_THEMES } from './theme-gallery-legacy';
import { THEME_EFFECT_PRESETS } from './theme-effect-presets';
import { findThemePreset } from './theme-presets';

export * from './theme-presets';
export { THEME_EFFECT_PRESETS, sameThemeEffectStyles } from './theme-effect-presets';

const office2007 = findThemePreset('office-2007-2010');

/** Featured themes followed by alphabetized themes from earlier Word releases. */
export const DOCUMENT_THEME_CATALOGUE = [
	...FEATURED_THEMES,
	...(office2007 ? [office2007] : []),
	...ADDITIONAL_THEMES,
].map((theme) => {
	const effect = THEME_EFFECT_PRESETS[theme.id === 'office-2007-2010' ? 1 : 0];
	return { ...theme, formatScheme: themeMatrix(effect?.name, effect?.styles) };
});
