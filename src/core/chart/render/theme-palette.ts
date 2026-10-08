// Theme slots in the SpreadsheetML index order every chart painter reads, with the Office
// defaults for slots a theme lacks.
import { parseArgbHex } from '../../color/hex-rgb';
import type { ChartThemePalette } from './summary';

/** Theme slot names in SpreadsheetML index order. */
export const THEME_SLOTS = [
	'lt1',
	'dk1',
	'lt2',
	'dk2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const;

/** Office 2013-2022 theme colours, used for slots a palette lacks. */
export const DEFAULT_THEME_COLORS: readonly string[] = [
	'FFFFFF',
	'000000',
	'E7E6E6',
	'44546A',
	'4472C4',
	'ED7D31',
	'A5A5A5',
	'FFC000',
	'5B9BD5',
	'70AD47',
	'0563C1',
	'954F72',
];

/** The `RRGGBB` of theme slot `index` (0 lt1, 1 dk1, 2 lt2, 3 dk2, 4-9 accents, 10 hlink, 11 folHlink). */
export function themeColor(theme: ChartThemePalette, index: number): string | undefined {
	const hex = theme.colors[index] ?? DEFAULT_THEME_COLORS[index];
	return hex && parseArgbHex(hex) ? hex.replace(/^#/, '') : undefined;
}
