export const THEME_EFFECT_NAMES = [
	'Office',
	'Office 2007-2010',
	'Subtle Solids',
	'Banded Edge',
	'Smokey Glass',
	'Glow Edge',
	'Frosted Glass',
	'Top Shadow',
	'Inset',
	'Milk Glass',
	'Riblet',
	'Glossy',
	'Reflection',
	'Extreme Shadow',
] as const;

export type ThemeEffectPresetName = (typeof THEME_EFFECT_NAMES)[number];
