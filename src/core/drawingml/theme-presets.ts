// Built-in Office theme presets, colour schemes and font schemes.
// Colours are stored in SpreadsheetML palette slot order:
// [lt1, dk1, lt2, dk2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink].
import { OFFICE_COLORS } from './office-theme-colors';
import {
	OFFICE_APTOS_COLORS,
	OFFICE_2007_COLORS,
	GRAYSCALE_COLORS,
	FACET_COLORS,
	INTEGRAL_COLORS,
	ION_COLORS,
	ORGANIC_COLORS,
	RETROSPECT_COLORS,
	SLICE_COLORS,
	BLUE_WARM_COLORS,
	GREEN_COLORS,
} from './theme-color-presets';

export type ThemeColorTuple = readonly [
	string,
	string,
	string,
	string,
	string,
	string,
	string,
	string,
	string,
	string,
	string,
	string,
];

export interface ThemeColorSchemePreset {
	id: string;
	label: string;
	colors: ThemeColorTuple;
}

export interface ThemeFontSchemePreset {
	id: string;
	label: string;
	majorFont: string;
	minorFont: string;
}

export interface ThemePreset {
	id: string;
	label: string;
	colors: ThemeColorTuple;
	majorFont: string;
	minorFont: string;
}

/** Built-in Office theme colour schemes for document theme galleries. */
export const THEME_COLOR_PRESETS: readonly ThemeColorSchemePreset[] = [
	{ id: 'office', label: 'Office', colors: OFFICE_COLORS },
	{ id: 'office-aptos', label: 'Office (Aptos)', colors: OFFICE_APTOS_COLORS },
	{ id: 'office-2007-2010', label: 'Office 2007-2010', colors: OFFICE_2007_COLORS },
	{ id: 'grayscale', label: 'Grayscale', colors: GRAYSCALE_COLORS },
	{ id: 'facet', label: 'Facet', colors: FACET_COLORS },
	{ id: 'integral', label: 'Integral', colors: INTEGRAL_COLORS },
	{ id: 'ion', label: 'Ion', colors: ION_COLORS },
	{ id: 'organic', label: 'Organic', colors: ORGANIC_COLORS },
	{ id: 'retrospect', label: 'Retrospect', colors: RETROSPECT_COLORS },
	{ id: 'slice', label: 'Slice', colors: SLICE_COLORS },
	{ id: 'blue-warm', label: 'Blue Warm', colors: BLUE_WARM_COLORS },
	{ id: 'green', label: 'Green', colors: GREEN_COLORS },
];

/** Built-in Office theme font schemes for document theme galleries. */
export const THEME_FONT_PRESETS: readonly ThemeFontSchemePreset[] = [
	{ id: 'office', label: 'Office', majorFont: 'Calibri Light', minorFont: 'Calibri' },
	{
		id: 'office-aptos',
		label: 'Office (Aptos)',
		majorFont: 'Aptos Display',
		minorFont: 'Aptos Narrow',
	},
	{ id: 'office-2007-2010', label: 'Office 2007-2010', majorFont: 'Cambria', minorFont: 'Calibri' },
	{ id: 'arial', label: 'Arial', majorFont: 'Arial', minorFont: 'Arial' },
	{ id: 'calibri', label: 'Calibri', majorFont: 'Calibri', minorFont: 'Calibri' },
	{ id: 'candara', label: 'Candara', majorFont: 'Candara', minorFont: 'Candara' },
	{
		id: 'century-gothic',
		label: 'Century Gothic',
		majorFont: 'Century Gothic',
		minorFont: 'Century Gothic',
	},
	{ id: 'consolas', label: 'Consolas', majorFont: 'Consolas', minorFont: 'Consolas' },
	{ id: 'corbel', label: 'Corbel', majorFont: 'Corbel', minorFont: 'Corbel' },
	{
		id: 'franklin-gothic',
		label: 'Franklin Gothic',
		majorFont: 'Franklin Gothic Medium',
		minorFont: 'Franklin Gothic Book',
	},
	{ id: 'garamond', label: 'Garamond', majorFont: 'Garamond', minorFont: 'Garamond' },
	{ id: 'georgia', label: 'Georgia', majorFont: 'Georgia', minorFont: 'Georgia' },
	{ id: 'segoe-ui', label: 'Segoe UI', majorFont: 'Segoe UI Light', minorFont: 'Segoe UI' },
	{
		id: 'trebuchet-ms',
		label: 'Trebuchet MS',
		majorFont: 'Trebuchet MS',
		minorFont: 'Trebuchet MS',
	},
	{ id: 'verdana', label: 'Verdana', majorFont: 'Verdana', minorFont: 'Verdana' },
];

/** Complete built-in Office themes (colours + fonts) for document theme galleries. */
export const THEME_PRESETS: readonly ThemePreset[] = [
	{
		id: 'office',
		label: 'Office',
		colors: OFFICE_COLORS,
		majorFont: 'Calibri Light',
		minorFont: 'Calibri',
	},
	{
		id: 'office-aptos',
		label: 'Office (Aptos)',
		colors: OFFICE_APTOS_COLORS,
		majorFont: 'Aptos Display',
		minorFont: 'Aptos Narrow',
	},
	{
		id: 'office-2007-2010',
		label: 'Office 2007-2010',
		colors: OFFICE_2007_COLORS,
		majorFont: 'Cambria',
		minorFont: 'Calibri',
	},
	{
		id: 'facet',
		label: 'Facet',
		colors: FACET_COLORS,
		majorFont: 'Trebuchet MS',
		minorFont: 'Trebuchet MS',
	},
	{
		id: 'integral',
		label: 'Integral',
		colors: INTEGRAL_COLORS,
		majorFont: 'Tw Cen MT',
		minorFont: 'Tw Cen MT',
	},
	{
		id: 'ion',
		label: 'Ion',
		colors: ION_COLORS,
		majorFont: 'Century Gothic',
		minorFont: 'Century Gothic',
	},
	{
		id: 'organic',
		label: 'Organic',
		colors: ORGANIC_COLORS,
		majorFont: 'Garamond',
		minorFont: 'Garamond',
	},
	{
		id: 'retrospect',
		label: 'Retrospect',
		colors: RETROSPECT_COLORS,
		majorFont: 'Calibri Light',
		minorFont: 'Calibri',
	},
	{
		id: 'slice',
		label: 'Slice',
		colors: SLICE_COLORS,
		majorFont: 'Century Gothic',
		minorFont: 'Century Gothic',
	},
];

export const findThemePreset = (id: string): ThemePreset | undefined =>
	THEME_PRESETS.find((preset) => preset.id === id);

export const findThemeColorPreset = (id: string): ThemeColorSchemePreset | undefined =>
	THEME_COLOR_PRESETS.find((preset) => preset.id === id);

export const findThemeFontPreset = (id: string): ThemeFontSchemePreset | undefined =>
	THEME_FONT_PRESETS.find((preset) => preset.id === id);
