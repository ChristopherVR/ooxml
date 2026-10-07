/**
 * Office's Chart Design > Change Colors palettes, resolved under a host theme.
 *
 * Ground truth: `viewers/pptx/scripts/capture-data-galleries-com.ps1` sets
 * `Chart.ChartColor = 10..26` on a clustered column chart and saves; each
 * chart's `colorsN.xml` is the `cs:colorStyle` below (id, `meth`, base scheme
 * colours, `cs:variation` list) and each `c:ser/c:spPr` the colour PowerPoint
 * derived from it. Two derivations exist:
 *
 * - `cycle`: series i takes base colour `i mod len`, transformed by variation
 *   `floor(i / len)` (variation 0 of the colourful palettes is "none").
 * - `withinLinear` (and `withinLinearReversed`): one base colour spread from a
 *   shade to a tint. Measured for 2..10 series (ChartColor 14 over 2, 3, 4,
 *   5, 6, 7, 8 and 10 series):
 *   the step is `140 / (n + 1)` percent, series at distance d from the centre
 *   get `100 - step * d`, written as `a:shade` (rounded down) before the
 *   centre and `a:tint` (rounded up) after it, the centre (odd n) untouched.
 *   Four series give shade 58000, shade 86000, tint 86000, tint 58000.
 *
 * Also checked against Excel 16.0 build 20430 through COM for all 17 palettes
 * with 1, 2, 4, 7, 10, 19 and 55 series under Office and custom themes
 * (`scripts/record-xlsx-chart-colors.ps1`).
 * The Office baseline up to 10 series matches exactly; extended probes expose
 * one-channel-step rounding differences in some colors. This catalog does not
 * establish exact native RGB parity for every theme or series count.
 *
 * @module chart/color-palettes
 */
import { resolveDrawingColor } from '../diagram/drawing-color';

export type ChartThemeColor =
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
export type ChartColorScheme = Readonly<Record<ChartThemeColor, string>>;

type Transform = Readonly<Partial<Record<'lumMod' | 'lumOff' | 'tint' | 'shade', number>>>;

export interface ChartColorPalette {
	/** `Chart.ChartColor` / `cs:colorStyle/@id`. */
	id: number;
	group: 'colorful' | 'monochromatic';
	/** 1-based number within its group ("Colorful Palette 3"). */
	index: number;
	meth: 'cycle' | 'withinLinear' | 'withinLinearReversed';
	base: ReadonlyArray<ChartThemeColor>;
	variations: readonly Transform[];
}

/** The colourful palettes' shared variation list (colors1.xml, id 10). */
const COLORFUL_VARIATIONS: readonly Transform[] = [
	{},
	{ lumMod: 60000 },
	{ lumMod: 80000, lumOff: 20000 },
	{ lumMod: 80000 },
	{ lumMod: 60000, lumOff: 40000 },
	{ lumMod: 50000 },
	{ lumMod: 70000, lumOff: 30000 },
	{ lumMod: 70000 },
	{ lumMod: 50000, lumOff: 50000 },
];

/** id 20's dk1 variation list (colors11.xml). */
const GRAY_VARIATIONS: readonly Transform[] = [
	{ tint: 88500 },
	{ tint: 55000 },
	{ tint: 75000 },
	{ tint: 98500 },
	{ tint: 30000 },
	{ tint: 60000 },
	{ tint: 80000 },
];

const ACCENTS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'] as const;

function colorful(id: number, index: number, base: ChartColorPalette['base']): ChartColorPalette {
	return { id, group: 'colorful', index, meth: 'cycle', base, variations: COLORFUL_VARIATIONS };
}

function mono(
	id: number,
	index: number,
	accent: ChartThemeColor,
	meth: ChartColorPalette['meth'],
): ChartColorPalette {
	return { id, group: 'monochromatic', index, meth, base: [accent], variations: [] };
}

export const CHART_COLOR_PALETTES: readonly ChartColorPalette[] = [
	colorful(10, 1, ACCENTS),
	colorful(11, 2, ['accent1', 'accent3', 'accent5']),
	colorful(12, 3, ['accent2', 'accent4', 'accent6']),
	colorful(13, 4, ['accent6', 'accent5', 'accent4']),
	...ACCENTS.map((accent, i) => mono(14 + i, 1 + i, accent, 'withinLinear')),
	{
		id: 20,
		group: 'monochromatic',
		index: 7,
		meth: 'cycle',
		base: ['dk1'],
		variations: GRAY_VARIATIONS,
	},
	...ACCENTS.map((accent, i) => mono(21 + i, 8 + i, accent, 'withinLinearReversed')),
];

export function findChartColorPalette(id: number): ChartColorPalette | undefined {
	return CHART_COLOR_PALETTES.find((palette) => palette.id === id);
}

function resolve(scheme: ChartColorScheme, key: ChartThemeColor, transform: Transform): string {
	return (
		resolveDrawingColor(
			{
				kind: 'scheme',
				value: key,
				transforms: Object.entries(transform).map(([name, value]) => ({
					name,
					value: String(value),
				})),
			},
			{ scheme: (name) => scheme[name as ChartThemeColor] },
		)?.hex ?? scheme[key]
	).toUpperCase();
}
/** The `a:shade` / `a:tint` PowerPoint writes for series `i` of `n` (withinLinear order). */
export function withinLinearTransform(i: number, n: number): Transform {
	if (n <= 1) {
		return {};
	}
	const step = 140 / (n + 1);
	const centre = (n - 1) / 2;
	const value = 100 - step * Math.abs(i - centre);
	if (i < centre) {
		return { shade: Math.floor(value + 1e-9) * 1000 };
	}
	if (i > centre) {
		return { tint: Math.ceil(value - 1e-9) * 1000 };
	}
	return {};
}

/** The colour of series `i` of `n` under `palette`, in the deck's theme. */
export function chartPaletteSeriesColor(
	palette: ChartColorPalette,
	i: number,
	n: number,
	scheme: ChartColorScheme,
): string {
	if (palette.meth === 'cycle') {
		const base = palette.base[i % palette.base.length];
		const round = Math.floor(i / palette.base.length);
		const transform = palette.variations.length
			? palette.variations[round % palette.variations.length]
			: {};
		if (!base) throw new RangeError('A chart palette must have a base color');
		return resolve(scheme, base, transform ?? {});
	}
	const index = palette.meth === 'withinLinearReversed' ? n - 1 - i : i;
	const base = palette.base[0];
	if (!base) throw new RangeError('A chart palette must have a base color');
	return resolve(scheme, base, withinLinearTransform(index, n));
}

/** The first `n` colours of `palette`. */
export function chartPaletteColors(
	palette: ChartColorPalette,
	n: number,
	scheme: ChartColorScheme,
): string[] {
	return Array.from({ length: n }, (_, i) => chartPaletteSeriesColor(palette, i, n, scheme));
}
