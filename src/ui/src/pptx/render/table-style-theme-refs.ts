/**
 * Resolve a table style's theme style-matrix references for rendering.
 *
 * "Themed Style 1" and "Themed Style 2" paint the table background through
 * `a:tblBg/a:fillRef` and draw their gridlines through `a:lnRef`, so their
 * look lives in the theme's format scheme (`a:fillStyleLst`, `a:lnStyleLst`)
 * rather than in the style. The loader resolves a deck-authored `tblBg`
 * reference up front, but the built-in catalogue cannot (it has no theme),
 * and no `lnRef` border was resolved at all: such a table drew no background
 * and no lines, while PowerPoint draws both even with every Table Style
 * Option off. This module resolves both against the deck's format scheme.
 *
 * The `tblBg` fill is painted as the lowest layer of each cell, as a
 * deck-authored `tblBg` already was. For a gradient fill style that is an
 * approximation: PowerPoint paints one gradient across the whole table, the
 * renderer one per cell.
 *
 * @module render/table-style-theme-refs
 */
import { resolveTableBackgroundFillRef, resolveTableStyleLineRef } from 'ooxml-core/pptx';
import type {
	ParsedTableStyleBorder,
	ParsedTableStyleBorders,
	ParsedTableStyleEntry,
	PptxThemeColorScheme,
	PptxThemeFormatScheme,
} from 'ooxml-core/pptx';

import { SCHEME_ALIASES } from './table-style-scheme';

/** The section border keys of a {@link ParsedTableStyleEntry}. */
const BORDER_KEYS = [
	'wholeTblBorders',
	'band1HBorders',
	'band2HBorders',
	'band1VBorders',
	'band2VBorders',
	'lastColBorders',
	'firstColBorders',
	'lastRowBorders',
	'seCellBorders',
	'swCellBorders',
	'firstRowBorders',
	'neCellBorders',
	'nwCellBorders',
] as const;

/** True when a border is only an unresolved `a:lnRef`. */
export function isLineRefOnly(border: ParsedTableStyleBorder | undefined): boolean {
	return Boolean(
		border?.lnRef &&
		border.width === undefined &&
		border.dash === undefined &&
		border.fill === undefined &&
		border.color === undefined &&
		border.noFill === undefined,
	);
}

/** A scheme key -> hex map from a colour scheme, with the default clrMap aliases. */
function colorMapOf(colorScheme: PptxThemeColorScheme | undefined): Record<string, string> {
	const map: Record<string, string> = {};
	for (const [key, value] of Object.entries(colorScheme ?? {})) {
		if (typeof value === 'string' && value) {
			map[key] = value;
		}
	}
	for (const [alias, key] of Object.entries(SCHEME_ALIASES)) {
		const value = map[key];
		if (!map[alias] && value) {
			map[alias] = value;
		}
	}
	return map;
}

function resolveBorders(
	borders: ParsedTableStyleBorders,
	formatScheme: PptxThemeFormatScheme | undefined,
	colorMap: Record<string, string>,
): ParsedTableStyleBorders {
	const next: ParsedTableStyleBorders = {};
	for (const [side, border] of Object.entries(borders) as [
		keyof ParsedTableStyleBorders,
		ParsedTableStyleBorder,
	][]) {
		if (!isLineRefOnly(border)) {
			next[side] = border;
			continue;
		}
		// An unresolvable reference draws nothing, as before it was parsed.
		const resolved = border.lnRef
			? resolveTableStyleLineRef(border.lnRef, formatScheme, colorMap)
			: undefined;
		if (resolved) {
			next[side] = resolved;
		}
	}
	return next;
}

/** True when `entry` has a `tblBg` fill reference or an `a:lnRef` border to resolve. */
function hasThemeRefs(entry: ParsedTableStyleEntry): boolean {
	const background = entry.tableBackground;
	if (background?.fillRef && !background.fill) {
		return true;
	}
	return BORDER_KEYS.some((key) => Object.values(entry[key] ?? {}).some(isLineRefOnly));
}

/**
 * `entry` with its `tblBg` fill reference and `lnRef` borders resolved
 * against `formatScheme` and `colorScheme`. References that cannot be
 * resolved (no format scheme, or no style at that index) are dropped. An
 * entry without references is returned as is.
 *
 * Nothing is cached by entry: the table style editor
 * (`applyTableStyleFieldEdit`) and the `set_table_style_section` tool edit an
 * entry in place, so a copy kept for an entry would go on painting the
 * pre-edit style.
 */
export function resolveTableStyleThemeRefs(
	entry: ParsedTableStyleEntry | undefined,
	formatScheme: PptxThemeFormatScheme | undefined,
	colorScheme: PptxThemeColorScheme | undefined,
): ParsedTableStyleEntry | undefined {
	if (!entry || !hasThemeRefs(entry)) {
		return entry;
	}
	const colorMap = colorMapOf(colorScheme);
	const next: ParsedTableStyleEntry = { ...entry };
	const background = entry.tableBackground;
	if (background?.fillRef && !background.fill) {
		const fill = resolveTableBackgroundFillRef(background.fillRef, formatScheme, colorMap);
		if (fill) {
			next.tableBackground = { ...background, fill };
		}
	}
	for (const key of BORDER_KEYS) {
		const borders = entry[key];
		if (borders) {
			next[key] = resolveBorders(borders, formatScheme, colorMap);
		}
	}
	return next;
}
