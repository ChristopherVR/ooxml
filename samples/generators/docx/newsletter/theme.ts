/** The newsletter's theme: the garden's palette as the colour scheme, Rubik and Noto Sans as fonts. */
import type { ThemeCatalog } from '../../../../src/core/docx/index';
import { FONTS } from '../../lib/fiction';

/** The garden's palette, as the theme's colour scheme. */
export const GARDEN = {
	ink: '2B2A26',
	white: 'FFFFFF',
	forest: '2F5D3A',
	cream: 'F6F1E4',
	leaf: '4C8B4A',
	terracotta: 'C8553D',
	marigold: 'E3B23C',
	larkspur: '6A4C93',
	sky: '3D7EA6',
	bark: '8C6E4A',
} as const;

export const THEME: ThemeCatalog = {
	name: 'Larkspur Garden',
	colorSchemeName: 'Larkspur Garden',
	fontSchemeName: 'Larkspur Garden',
	colors: {
		dk1: GARDEN.ink,
		lt1: GARDEN.white,
		dk2: GARDEN.forest,
		lt2: GARDEN.cream,
		accent1: GARDEN.leaf,
		accent2: GARDEN.terracotta,
		accent3: GARDEN.marigold,
		accent4: GARDEN.larkspur,
		accent5: GARDEN.sky,
		accent6: GARDEN.bark,
		hlink: GARDEN.forest,
		folHlink: GARDEN.larkspur,
	},
	colorMapping: {},
	fonts: {
		major: { latin: FONTS.display, eastAsia: FONTS.body, complexScript: FONTS.body },
		minor: { latin: FONTS.body, eastAsia: FONTS.body, complexScript: FONTS.body },
	},
};
