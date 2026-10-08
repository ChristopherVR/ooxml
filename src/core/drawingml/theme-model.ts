// Format-neutral DrawingML theme model (ECMA-376 Part 1, 20.1.6): the colour scheme, the font
// scheme and the colour map that ties the logical bg/tx names to scheme slots. Every Office format
// carries the same `a:theme` part (`word/theme`, `xl/theme`, `ppt/theme`, Visio's theme).
import type { DrawingColor } from './types';

/** The twelve `a:clrScheme` slots, in schema order. */
export type ThemeColorSlot =
	| 'dk1'
	| 'lt1'
	| 'dk2'
	| 'lt2'
	| 'accent1'
	| 'accent2'
	| 'accent3'
	| 'accent4'
	| 'accent5'
	| 'accent6'
	| 'hlink'
	| 'folHlink';

/** The colour scheme slots in schema order. */
export const THEME_COLOR_SLOTS: readonly ThemeColorSlot[] = [
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
];

/** The four logical background and text names. */
export type ThemeLogicalColor = 'bg1' | 'tx1' | 'bg2' | 'tx2';

/** The names a colour map (`p:clrMap`, `a:overrideClrMapping`) assigns to scheme slots. */
export type ThemeColorMapKey =
	| ThemeLogicalColor
	| Exclude<ThemeColorSlot, 'dk1' | 'lt1' | 'dk2' | 'lt2'>;

/** A colour map: which scheme slot each name uses. Missing keys use `DEFAULT_THEME_COLOR_MAP`. */
export type ThemeColorMap = Partial<Record<ThemeColorMapKey, ThemeColorSlot>>;

/** The conventional mapping (light background, dark text) every Office default master uses. */
export const DEFAULT_THEME_COLOR_MAP: Readonly<Record<ThemeLogicalColor, ThemeColorSlot>> = {
	bg1: 'lt1',
	tx1: 'dk1',
	bg2: 'lt2',
	tx2: 'dk2',
};

/** A parsed `a:clrScheme`. Colours stay as written (a `sysClr` keeps its `lastClr` as `fallback`). */
export interface ThemeColorScheme {
	name?: string;
	colors: Partial<Record<ThemeColorSlot, DrawingColor>>;
}

/** One `a:majorFont` or `a:minorFont`. Empty typefaces are omitted. */
export interface ThemeFontCollection {
	latin?: string;
	eastAsia?: string;
	complexScript?: string;
	/** `a:font` overrides by script tag (`Jpan`, `Arab`...). */
	scripts: Record<string, string>;
}

/** A parsed `a:fontScheme`. */
export interface ThemeFontScheme {
	name?: string;
	major: ThemeFontCollection;
	minor: ThemeFontCollection;
}

/** A parsed `a:theme` (or `a:themeOverride`): the parts every format shares. */
export interface DrawingTheme {
	/** `a:theme/@name`. */
	name?: string;
	colorScheme: ThemeColorScheme;
	fontScheme: ThemeFontScheme;
}
