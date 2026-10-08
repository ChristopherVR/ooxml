/**
 * chart-style-palette.ts: the series colour cycle of a chart with no explicit
 * series colours.
 *
 * Precedence, decided once here for every binding ({@link resolveChartPalette}):
 * the palette parsed from the chart's colour-style part (`colorsN.xml`, already
 * resolved against the theme by core) wins; otherwise the built-in chart style
 * (`c:style/@val`) picks a column of Office's chart-style gallery, built over
 * the deck theme's accent colours when core resolved them
 * (`PptxChartData.themeAccentColors`), or over the Office theme's when it did not.
 *
 * @module chart-style-palette
 */
import type { PptxChartData } from 'ooxml-core/pptx';

/** Parse a hex colour string (#RRGGBB) into [r, g, b]. */
function hexToRgb(hex: string): [number, number, number] {
	const h = hex.replace('#', '');
	return [
		parseInt(h.substring(0, 2), 16),
		parseInt(h.substring(2, 4), 16),
		parseInt(h.substring(4, 6), 16),
	];
}

/** Convert [r, g, b] back to #RRGGBB. */
function rgbToHex(r: number, g: number, b: number): string {
	const clamp = (v: number): number => Math.max(0, Math.min(255, Math.round(v)));
	return `#${clamp(r).toString(16).padStart(2, '0')}${clamp(g)
		.toString(16)
		.padStart(2, '0')}${clamp(b).toString(16).padStart(2, '0')}`;
}

/**
 * Apply a tint (lighten towards white) to a colour.
 * `amount` in [0, 1] where 0 = no change, 1 = white.
 */
export function tint(hex: string, amount: number): string {
	const [r, g, b] = hexToRgb(hex);
	return rgbToHex(r + (255 - r) * amount, g + (255 - g) * amount, b + (255 - b) * amount);
}

/**
 * Apply a shade (darken towards black) to a colour.
 * `amount` in [0, 1] where 0 = no change, 1 = black.
 */
export function shade(hex: string, amount: number): string {
	const [r, g, b] = hexToRgb(hex);
	return rgbToHex(r * (1 - amount), g * (1 - amount), b * (1 - amount));
}

/** The Office theme's accent1 to accent6, used when a chart carries no theme accents. */
const OFFICE_ACCENTS: readonly string[] = [
	'#4472C4',
	'#ED7D31',
	'#A5A5A5',
	'#FFC000',
	'#5B9BD5',
	'#70AD47',
];

/** The two colours after the six accents in {@link DEFAULT_CHART_PALETTE}. */
const DEFAULT_EXTRAS: readonly string[] = ['#FF0000', '#00B0F0'];

/**
 * The default fallback palette used when no chart style is specified.
 *
 * The Office accent cycle (accent1-6 plus the two chart extras), i.e. what
 * PowerPoint paints for a chart with no style part. This MUST stay identical
 * to `chart-view-model.ts`'s `DEFAULT_PALETTE` (which aliases it): the two
 * used to differ (this one was a Tailwind-ish set), and since the bindings
 * reach the shared engine through both entry points, the same deck fell back
 * to two different palettes depending on the binding.
 */
export const DEFAULT_CHART_PALETTE: ReadonlyArray<string> = [...OFFICE_ACCENTS, ...DEFAULT_EXTRAS];

/** Monochrome ramp of one colour, darkest first: three shades, the colour, four tints. */
function monochromaticRamp(base: string): string[] {
	return [
		shade(base, 0.5),
		shade(base, 0.35),
		shade(base, 0.15),
		base,
		tint(base, 0.2),
		tint(base, 0.4),
		tint(base, 0.6),
		tint(base, 0.8),
	];
}

/** The greyscale column's base (text 1 at 50%). */
const GREY = '#7F7F7F';

/**
 * Office's built-in chart styles (`c:style/@val`, 1-48). ECMA-376 Part 1, 21.2.2.196 only says the
 * value selects one of the application's predefined styles; the colours follow Office's chart style
 * gallery of six rows of eight: column `(id - 1) % 8` is 0 greyscale, 1 colourful (accent1 to
 * accent6 in theme order, then the same accents darker for series 7-12: Office darkens them with
 * `lumMod`, an RGB shade of 40% approximates it; style 2 is Office's default), and 2 to 7
 * monochrome shades and tints of accent1 to accent6. The rows (1-8, 9-16, 17-24, 25-32, 33-40,
 * 41-48) vary outlines, effects and backgrounds, not the series colours.
 */
function buildPalette(column: number, accents: readonly string[]): string[] {
	if (column === 0) {
		return monochromaticRamp(GREY);
	}
	if (column === 1) {
		return [...accents, ...accents.map((accent) => shade(accent, 0.4))];
	}
	return monochromaticRamp(accents[column - 2]!);
}

/** `accents` when it is six `#RRGGBB` colours, else the Office theme's. */
function usableAccents(accents: readonly string[] | undefined): readonly string[] {
	return accents?.length === 6 && accents.every((accent) => /^#[0-9a-f]{6}$/iu.test(accent))
		? accents
		: OFFICE_ACCENTS;
}

const paletteCache = new Map<string, readonly string[]>();

/**
 * Get the colour palette for a chart style index (1–48), built over
 * `themeAccents` (the deck theme's accent1 to accent6; the Office theme's when
 * absent). Falls back to {@link DEFAULT_CHART_PALETTE}, with the theme's accents
 * in place of the Office ones, when `styleId` is undefined or out of range.
 */
export function getChartStylePalette(
	styleId?: number,
	themeAccents?: readonly string[],
): ReadonlyArray<string> {
	const accents = usableAccents(themeAccents);
	const valid = styleId !== undefined && styleId >= 1 && styleId <= 48;
	if (!valid && accents === OFFICE_ACCENTS) {
		return DEFAULT_CHART_PALETTE;
	}
	const column = valid ? (Math.trunc(styleId) - 1) % 8 : -1;
	const key = `${column}:${accents.join(',')}`;
	let palette = paletteCache.get(key);
	if (!palette) {
		palette = column < 0 ? [...accents, ...DEFAULT_EXTRAS] : buildPalette(column, accents);
		paletteCache.set(key, palette);
	}
	return palette;
}

/**
 * The series colour cycle a chart paints with: its parsed colour-style palette
 * when it has one, otherwise its chart style's palette over the deck theme.
 * Every binding resolves the palette through here (and `buildChartViewModel`
 * applies it), so the same chart cannot paint two colour cycles.
 */
export function resolveChartPalette(chartData: PptxChartData): string[] {
	if (chartData.colorPalette && chartData.colorPalette.length > 0) {
		return [...chartData.colorPalette];
	}
	return [...getChartStylePalette(chartData.style?.styleId, chartData.themeAccentColors)];
}

/** `chartData` with {@link resolveChartPalette}'s palette as its `colorPalette`. */
export function withResolvedPalette(chartData: PptxChartData): PptxChartData {
	return chartData.colorPalette && chartData.colorPalette.length > 0
		? chartData
		: { ...chartData, colorPalette: resolveChartPalette(chartData) };
}
