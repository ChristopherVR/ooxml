// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import { parseTheme as parseDrawingTheme } from '../drawingml/theme';
import { THEME_COLOR_SLOTS, type ThemeFontCollection } from '../drawingml/theme-model';
import { themeSlotHex } from '../drawingml/theme-color';
import { parseXml, WORD_NS, getW, type XmlElement } from './xml';
import type { ThemeCatalog, ThemeColorSlot, ThemeFontSet } from './theme-model';
import { isStWmlColorSchemeIndex, type StWmlColorSchemeIndex } from './generated/wml-simple-types';
import { enumValue } from './parse-diagnostics';

function fontSetOf(collection: ThemeFontCollection): ThemeFontSet {
	const { latin, eastAsia, complexScript } = collection;
	return {
		...(latin ? { latin } : {}),
		...(eastAsia ? { eastAsia } : {}),
		...(complexScript ? { complexScript } : {}),
	};
}

/**
 * Parses `word/theme/theme1.xml`'s color scheme and major/minor font scheme through the neutral
 * DrawingML theme parser. Scheme colours keep the `RRGGBB` of an `srgbClr` or a `sysClr`'s
 * `lastClr`; other colour kinds leave the slot empty.
 */
export function parseTheme(xml: string): ThemeCatalog {
	const theme = parseDrawingTheme(parseXml(xml));
	const colors: Partial<Record<ThemeColorSlot, string>> = {};
	for (const slot of THEME_COLOR_SLOTS) {
		const value = themeSlotHex(theme.colorScheme.colors[slot]);
		if (value) colors[slot] = value;
	}
	const { major, minor } = theme.fontScheme;
	return { colors, colorMapping: {}, fonts: { major: fontSetOf(major), minor: fontSetOf(minor) } };
}

const MAPPING_KEYS = ['bg1', 'tx1', 'bg2', 'tx2'] as const;
const TOKEN_TO_SLOT: Record<StWmlColorSchemeIndex, ThemeColorSlot> = {
	dark1: 'dk1',
	light1: 'lt1',
	dark2: 'dk2',
	light2: 'lt2',
	accent1: 'accent1',
	accent2: 'accent2',
	accent3: 'accent3',
	accent4: 'accent4',
	accent5: 'accent5',
	accent6: 'accent6',
	hyperlink: 'hlink',
	followedHyperlink: 'folHlink',
};

/** Parses `word/settings.xml`'s `w:clrSchemeMapping`, mapping logical bg/tx slots to scheme slots. */
export function parseColorSchemeMapping(xml: string): ThemeCatalog['colorMapping'] {
	const document = parseXml(xml);
	const mapping = Array.from(document.getElementsByTagNameNS(WORD_NS, 'clrSchemeMapping'))[0] as
		| XmlElement
		| undefined;
	const result: ThemeCatalog['colorMapping'] = {};
	if (!mapping) return result;
	for (const key of MAPPING_KEYS) {
		const value = enumValue(
			isStWmlColorSchemeIndex,
			getW(mapping, key),
			`w:clrSchemeMapping/@w:${key}`,
		);
		if (value) result[key] = TOKEN_TO_SLOT[value];
	}
	return result;
}
