import { parseTheme as parseDrawingTheme } from '../../drawingml/theme';
import type { ThemeColorSlot } from '../../drawingml/theme-model';
import type { DrawingColor } from '../../drawingml/types';
import { parseXml } from '../../xml/index';
import type { ThemePalette } from '../model';
import { DEFAULT_THEME } from '../workbook';

/**
 * SpreadsheetML theme indices swap the first two pairs: `theme="0"` is lt1 (background 1) and
 * `theme="1"` is dk1 (text 1). The palette is stored in that index order.
 */
export const PALETTE_SLOTS = [
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
] as const satisfies readonly ThemeColorSlot[];

/** A scheme colour as `RRGGBB`: a `sysClr` reads `lastClr`, then `val`; other kinds read `val`. */
function paletteHex(color: DrawingColor | undefined): string | undefined {
	const value = color?.kind === 'system' ? (color.fallback ?? color.value) : color?.value;
	return value && /^[0-9A-Fa-f]{6}$/.test(value) ? value.toUpperCase() : undefined;
}

/** Reads the palette and the major/minor Latin fonts of a theme part (the neutral DrawingML theme). */
export function parseTheme(xml: string | undefined): ThemePalette {
	const palette: ThemePalette = { ...DEFAULT_THEME, colors: [...DEFAULT_THEME.colors] };
	if (!xml) return palette;
	const theme = parseDrawingTheme(parseXml(xml, { label: 'XLSX theme' }));
	palette.colors = PALETTE_SLOTS.map(
		(slot, index) =>
			paletteHex(theme.colorScheme.colors[slot]) ?? DEFAULT_THEME.colors[index] ?? '000000',
	);
	palette.majorFont = theme.fontScheme.major.latin ?? palette.majorFont;
	palette.minorFont = theme.fontScheme.minor.latin ?? palette.minorFont;
	return palette;
}
