/**
 * Design > Variants > Background Styles: the twelve theme background styles.
 *
 * PowerPoint offers a 4 by 3 grid. Each row is one entry of the theme's `a:bgFillStyleLst`
 * (`bgRef idx` 1001, 1002, 1003). Each column is a colour: the theme's light background (`bg1`),
 * its light alternate (`bg2`), then the dark pair, which PowerPoint writes as the same
 * `bgRef` with the colour map swapped (`bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"`), so text flips
 * to the light colour. Verified against PowerPoint's `Slide.BackgroundStyle` and
 * `SlideMaster.BackgroundStyle` output.
 *
 * @module pptx-types/background-style
 */

/** One cell of the Background Styles grid. */
export interface PptxBackgroundStylePreset {
	/** 1 to 12, row by row (PowerPoint's `msoBackgroundStylePreset1` to `12`). */
	index: number;
	/** `bgRef idx`: 1001, 1002 or 1003 into `a:bgFillStyleLst`. */
	idx: 1001 | 1002 | 1003;
	/** The scheme colour the fill style takes as `phClr`. */
	scheme: 'bg1' | 'bg2';
	/** True when the colour map is swapped to the dark pair (text turns light). */
	dark: boolean;
}

const COLUMNS: readonly { scheme: 'bg1' | 'bg2'; dark: boolean }[] = [
	{ scheme: 'bg1', dark: false },
	{ scheme: 'bg2', dark: false },
	{ scheme: 'bg2', dark: true },
	{ scheme: 'bg1', dark: true },
];

/** The twelve background styles in gallery order. */
export const BACKGROUND_STYLE_PRESETS: readonly PptxBackgroundStylePreset[] = [1001, 1002, 1003]
	.flatMap((idx) => COLUMNS.map((column) => ({ idx, ...column })))
	.map((entry, position) => ({
		index: position + 1,
		idx: entry.idx as 1001 | 1002 | 1003,
		scheme: entry.scheme,
		dark: entry.dark,
	}));

/** The preset at `index` (1 to 12), or `undefined`. */
export function backgroundStylePreset(index: number): PptxBackgroundStylePreset | undefined {
	return BACKGROUND_STYLE_PRESETS.find((preset) => preset.index === index);
}

/** The preset a `bgRef` (idx and scheme colour) and a dark flag describe, or `undefined`. */
export function backgroundStylePresetFor(
	idx: number,
	scheme: string,
	dark: boolean,
): PptxBackgroundStylePreset | undefined {
	return BACKGROUND_STYLE_PRESETS.find(
		(preset) => preset.idx === idx && preset.scheme === scheme && preset.dark === dark,
	);
}

const ACCENT_ALIASES = [
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const;

/** The twelve-entry colour map for the light (`false`) or dark (`true`) background pair. */
export function backgroundStyleColorMap(dark: boolean): Record<string, string> {
	const map: Record<string, string> = dark
		? { bg1: 'dk1', tx1: 'lt1', bg2: 'dk2', tx2: 'lt2' }
		: { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
	for (const alias of ACCENT_ALIASES) {
		map[alias] = alias;
	}
	return map;
}

/** True when a colour map (`bg1` routed to `dk1`) is the dark pair. */
export function isDarkColorMap(map: Readonly<Record<string, string>> | null | undefined): boolean {
	return map?.['bg1'] === 'dk1';
}

/**
 * A slide's theme background style, as the editor applied it or the loader read it.
 *
 * `kind: 'preset'` is `<p:bg><p:bgRef idx scheme/></p:bg>`; `kind: 'inherit'` is no `p:bg` at all
 * ("Follow master background"). The resolved values pin the reference to the flat
 * `backgroundColor`/`backgroundGradient` it produced: when a later edit (Format Background) moves
 * them the reference is stale and ignored on save.
 */
export interface PptxSlideBackgroundStyleRef {
	kind: 'preset' | 'inherit';
	/** `bgRef idx` (preset only). */
	idx?: number;
	/** `bgRef` scheme colour (preset only). */
	scheme?: 'bg1' | 'bg2';
	/** `backgroundColor` the style resolved to when it was set. */
	resolvedColor?: string;
	/** `backgroundGradient` the style resolved to when it was set. */
	resolvedGradient?: string;
}
