// Theme reads have implicit same-sheet inputs in addition to their argument references.
// Unsupported evaluation still rejects an affected theme cache; independent caches stay intact.
export const themeInputs = [
	'ColorSchemeIndex',
	'EffectSchemeIndex',
	'FontSchemeIndex',
	'ThemeIndex',
	'QuickStyleLineColor',
	'QuickStyleFillColor',
	'QuickStyleShadowColor',
	'QuickStyleFontColor',
	'QuickStyleLineMatrix',
	'QuickStyleFillMatrix',
	'QuickStyleEffectsMatrix',
	'QuickStyleFontMatrix',
	'QuickStyleVariation',
];
