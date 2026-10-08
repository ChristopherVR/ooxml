// DOM parser for the DrawingML theme part (`a:theme`) and its overrides (`a:themeOverride`), and
// for colour maps (`p:clrMap`, `a:overrideClrMapping`).
import { NS, elements, first, parseXml, type XmlDocument, type XmlElement } from '../xml/index';
import { parseDrawingColorIn } from './drawing-color';
import {
	THEME_COLOR_SLOTS,
	type DrawingTheme,
	type ThemeColorMap,
	type ThemeColorMapKey,
	type ThemeColorScheme,
	type ThemeColorSlot,
	type ThemeFontCollection,
	type ThemeFontScheme,
} from './theme-model';

const MAP_KEYS: readonly ThemeColorMapKey[] = [
	'bg1',
	'tx1',
	'bg2',
	'tx2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
];
const SLOT_SET: ReadonlySet<string> = new Set(THEME_COLOR_SLOTS);

/** Whether `value` names one of the twelve colour scheme slots. */
export function isThemeColorSlot(value: string): value is ThemeColorSlot {
	return SLOT_SET.has(value);
}

function nonEmpty(value: string | null | undefined): string | undefined {
	return value ? value : undefined;
}

/** Reads an `a:clrScheme`. */
export function parseThemeColorScheme(scheme: XmlElement | undefined): ThemeColorScheme {
	const result: ThemeColorScheme = { colors: {} };
	if (!scheme) return result;
	const name = nonEmpty(scheme.getAttribute('name'));
	if (name) result.name = name;
	for (const slot of THEME_COLOR_SLOTS) {
		const color = parseDrawingColorIn(first(scheme, slot, NS.a));
		if (color) result.colors[slot] = color;
	}
	return result;
}

function parseFontCollection(collection: XmlElement | undefined): ThemeFontCollection {
	const result: ThemeFontCollection = { scripts: {} };
	if (!collection) return result;
	const typeface = (local: string) =>
		nonEmpty(first(collection, local, NS.a)?.getAttribute('typeface'));
	const latin = typeface('latin');
	const eastAsia = typeface('ea');
	const complexScript = typeface('cs');
	if (latin) result.latin = latin;
	if (eastAsia) result.eastAsia = eastAsia;
	if (complexScript) result.complexScript = complexScript;
	for (const font of elements(collection)) {
		if (font.localName !== 'font') continue;
		const script = nonEmpty(font.getAttribute('script'));
		const face = nonEmpty(font.getAttribute('typeface'));
		if (script && face && !Object.hasOwn(result.scripts, script)) result.scripts[script] = face;
	}
	return result;
}

/** Reads an `a:fontScheme`. */
export function parseThemeFontScheme(scheme: XmlElement | undefined): ThemeFontScheme {
	const result: ThemeFontScheme = {
		major: parseFontCollection(first(scheme, 'majorFont', NS.a)),
		minor: parseFontCollection(first(scheme, 'minorFont', NS.a)),
	};
	const name = nonEmpty(scheme?.getAttribute('name'));
	if (name) result.name = name;
	return result;
}

function rootOf(source: XmlDocument | XmlElement): XmlElement {
	return 'documentElement' in source ? source.documentElement : source;
}

/**
 * Reads a theme part: `a:theme` (schemes under `a:themeElements`) or `a:themeOverride` (schemes
 * directly under the root). Missing schemes give empty ones; nothing is defaulted.
 */
export function parseTheme(source: XmlDocument | XmlElement): DrawingTheme {
	const root = rootOf(source);
	const holder = first(root, 'themeElements', NS.a) ?? root;
	const theme: DrawingTheme = {
		colorScheme: parseThemeColorScheme(first(holder, 'clrScheme', NS.a)),
		fontScheme: parseThemeFontScheme(first(holder, 'fontScheme', NS.a)),
	};
	const name = nonEmpty(root.getAttribute('name'));
	if (name) theme.name = name;
	return theme;
}

/** Parses theme part text; `label` names the format in parse errors (`XLSX theme`). */
export function parseThemeXml(xml: string, label = 'DrawingML theme'): DrawingTheme {
	return parseTheme(parseXml(xml, { label }));
}

/**
 * Reads a colour map element (`p:clrMap`, `a:overrideClrMapping`): each name to the scheme slot
 * it uses. Unknown names and slots are skipped.
 */
export function parseThemeColorMap(element: XmlElement | undefined): ThemeColorMap {
	const map: ThemeColorMap = {};
	if (!element) return map;
	for (const key of MAP_KEYS) {
		const value = element.getAttribute(key);
		if (value && isThemeColorSlot(value)) map[key] = value;
	}
	return map;
}
