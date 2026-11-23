// Built-in Office theme presets, colour schemes and font schemes.
// Colours are stored in SpreadsheetML palette slot order:
// [lt1, dk1, lt2, dk2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink].
import { DEFAULT_THEME } from './workbook';

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

const OFFICE_COLORS: ThemeColorTuple = [
	DEFAULT_THEME.colors[0] ?? 'FFFFFF',
	DEFAULT_THEME.colors[1] ?? '000000',
	DEFAULT_THEME.colors[2] ?? 'E7E6E6',
	DEFAULT_THEME.colors[3] ?? '44546A',
	DEFAULT_THEME.colors[4] ?? '4472C4',
	DEFAULT_THEME.colors[5] ?? 'ED7D31',
	DEFAULT_THEME.colors[6] ?? 'A5A5A5',
	DEFAULT_THEME.colors[7] ?? 'FFC000',
	DEFAULT_THEME.colors[8] ?? '5B9BD5',
	DEFAULT_THEME.colors[9] ?? '70AD47',
	DEFAULT_THEME.colors[10] ?? '0563C1',
	DEFAULT_THEME.colors[11] ?? '954F72',
];

const OFFICE_APTOS_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'E8E8E8',
	'0E2841',
	'156082',
	'E97132',
	'196B24',
	'0F9ED5',
	'A02B93',
	'4EA72E',
	'467886',
	'96607D',
];

const OFFICE_2007_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'EEECE1',
	'1F497D',
	'4F81BD',
	'C0504D',
	'9BBB59',
	'8064A2',
	'4BACC6',
	'F79646',
	'0000FF',
	'800080',
];

const GRAYSCALE_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'DDDDDD',
	'3F3F3F',
	'7F7F7F',
	'595959',
	'3F3F3F',
	'262626',
	'0D0D0D',
	'A5A5A5',
	'404040',
	'262626',
];

const FACET_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'DADEDF',
	'2C3E50',
	'90C226',
	'54A021',
	'E6B91E',
	'E76618',
	'C42F1A',
	'918655',
	'99CA3C',
	'B9D181',
];

const INTEGRAL_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'DFE3E5',
	'335B74',
	'1CADE4',
	'2683C6',
	'27CED7',
	'42BA97',
	'3E8853',
	'62A39F',
	'6BB76D',
	'B5D1CC',
];

const ION_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'E7E6E6',
	'44546A',
	'B01513',
	'EA6312',
	'E6B729',
	'6AAC90',
	'54849A',
	'9E5E9B',
	'58C1BA',
	'9D57A3',
];

const ORGANIC_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'E3DED1',
	'323232',
	'83992A',
	'3C9770',
	'44709D',
	'A23C33',
	'D97828',
	'DEB340',
	'69B764',
	'95B3D7',
];

const RETROSPECT_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'F3F2DC',
	'595959',
	'E48312',
	'BD582C',
	'865640',
	'9B8357',
	'C2BC80',
	'94A088',
	'E48312',
	'BD582C',
];

const SLICE_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'E3EBF1',
	'262626',
	'0563C1',
	'4BACC6',
	'2E8B57',
	'92D050',
	'FFC000',
	'ED7D31',
	'0563C1',
	'954F72',
];

const BLUE_WARM_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'EEEDF1',
	'242852',
	'4A66AC',
	'629DD1',
	'297FD5',
	'7F8FA9',
	'5AA2AE',
	'9D90A0',
	'9454C3',
	'3EBBF0',
];

const GREEN_COLORS: ThemeColorTuple = [
	'FFFFFF',
	'000000',
	'E8E7E1',
	'3B3838',
	'549E39',
	'8AB833',
	'C0CF3A',
	'029676',
	'4AB5C4',
	'0989B1',
	'82B917',
	'549E39',
];

/** Built-in Office theme colour schemes for Page Layout > Themes > Colors. */
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

/** Built-in Office theme font schemes for Page Layout > Themes > Fonts. */
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

/** Complete built-in Office themes (colours + fonts) for Page Layout > Themes. */
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
