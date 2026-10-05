import { NS, elements, first, parseXml } from '../../xml/index.js';
import type { ThemePalette } from '../model.js';
import { DEFAULT_THEME } from '../workbook.js';
import { att } from './xml-util.js';

/** Theme colour slots in `a:clrScheme` document order. */
export const SCHEME_ORDER = [
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const;

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
] as const;

/** Reads the palette and the major/minor Latin fonts of a theme part. */
export function parseTheme(xml: string | undefined): ThemePalette {
	const palette: ThemePalette = { ...DEFAULT_THEME, colors: [...DEFAULT_THEME.colors] };
	if (!xml) return palette;
	const root = parseXml(xml, { label: 'XLSX theme' }).documentElement;
	const elementsNode = first(root, 'themeElements', NS.a);
	const scheme = first(elementsNode, 'clrScheme', NS.a);
	if (scheme) {
		const bySlot = new Map<string, string>();
		for (const slot of elements(scheme)) {
			const colour = elements(slot)[0];
			if (!colour) continue;
			const value =
				colour.localName === 'sysClr'
					? (att(colour, 'lastClr') ?? att(colour, 'val'))
					: att(colour, 'val');
			if (value && /^[0-9A-Fa-f]{6}$/.test(value))
				bySlot.set(slot.localName ?? '', value.toUpperCase());
		}
		palette.colors = PALETTE_SLOTS.map(
			(slot, index) => bySlot.get(slot) ?? DEFAULT_THEME.colors[index] ?? '000000',
		);
	}
	const fonts = first(elementsNode, 'fontScheme', NS.a);
	const latin = (kind: string) =>
		att(first(first(fonts, kind, NS.a), 'latin', NS.a), 'typeface') || undefined;
	palette.majorFont = latin('majorFont') ?? palette.majorFont;
	palette.minorFont = latin('minorFont') ?? palette.minorFont;
	return palette;
}
