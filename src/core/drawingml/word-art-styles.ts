import type { WordArtStyle } from './word-art';

const accents = [
	['1769A5', '103D63', 'B7DDF5'],
	['B75A27', '763514', 'F6CEAE'],
	['1A785B', '0C4837', 'BCEBDD'],
	['783D99', '4D236A', 'E6CEF7'],
	['A73758', '681E39', 'F2BACC'],
	['71741C', '45460D', 'E4E9AB'],
] as const;

export const WORD_ART_STYLES: readonly WordArtStyle[] = [
	{ fill: '26374A' },
	{ fill: '1769A5', outline: '103D63' },
	{ fill: 'FFFFFF', outline: '323B57' },
	{ fill: 'B75A27', outline: '763514' },
	{ fill: '1A785B', outline: '0C4837' },
	{ fill: '783D99', outline: '4D236A' },
	...accents.flatMap(([fill, dark, light]) => [
		{ fill: light, outline: dark },
		{ fill: dark, outline: light },
		{ fill, shadow: dark },
		{ fill: 'FFFFFF', outline: dark, shadow: dark },
	]),
];
