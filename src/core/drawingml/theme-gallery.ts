import { themeMatrix } from './theme-matrix';
import { ADDITIONAL_THEMES, FEATURED_THEMES } from './theme-gallery-data';
import { THEME_EFFECT_PRESETS } from './theme-effect-presets';
import { findThemePreset } from './theme-presets';
import { OFFICE_APTOS_COLORS } from './theme-color-presets';
import { OFFICE_COLORS } from './office-theme-colors';

export * from './theme-presets';
export { THEME_EFFECT_PRESETS, sameThemeEffectStyles } from './theme-effect-presets';

const featuredOrder = [
	'facet',
	'gallery',
	'integral',
	'ion',
	'ion-boardroom',
	'organic',
	'retrospect',
	'slice',
	'wisp',
	'office-2007-2010',
];
const featured = featuredOrder.flatMap((id) => {
	const theme = FEATURED_THEMES.find((entry) => entry.id === id) ?? findThemePreset(id);
	return theme ? [theme] : [];
});

/** Featured themes followed by alphabetized themes from earlier Word releases. */
export const DOCUMENT_THEME_CATALOGUE = [
	{
		id: 'office',
		label: 'Office',
		colors: OFFICE_APTOS_COLORS,
		majorFont: 'Aptos Display',
		minorFont: 'Aptos',
	},
	...featured,
	{
		id: 'office-2013-2022',
		label: 'Office Theme 2013-2022',
		colors: OFFICE_COLORS,
		majorFont: 'Calibri Light',
		minorFont: 'Calibri',
	},
	...ADDITIONAL_THEMES,
].map((theme) => {
	const effect = THEME_EFFECT_PRESETS[theme.id === 'office-2007-2010' ? 1 : 0];
	return { ...theme, formatScheme: themeMatrix(effect?.name, effect?.styles) };
});
