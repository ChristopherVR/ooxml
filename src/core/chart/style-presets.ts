/**
 * The curated Chart Styles presets every product's Chart Design gallery offers (and PowerPoint's
 * on-canvas quick action): a one-click recolour of the whole chart with a built-in style number
 * mirrored onto `c:style`. Deliberately modest (six presets) rather than Office's per-chart-type
 * style gallery: `colorful` and `monochrome` are Office's style 2 and style 3 palettes; the four
 * variants are recolours of the colourful palette and mirror a colourful-column style of another
 * gallery row (Office's rows differ in effects, not colours, so no style number has these exact
 * colours). Office's 2013+ chart styles (`styleN.xml` parts, ids 201+) are not authored.
 *
 * Moved from `src/ui/src/pptx/render/chart-quick-action-styles.ts` (the preset table).
 *
 * @module chart/style-presets
 */
import { getChartStylePalette, shade, tint } from './style-palette';

export interface ChartStylePreset {
	id: string;
	/** English name. */
	label: string;
	/** The `c:style/@val` (1..48) mirrored onto the chart. */
	styleId: number;
	/** The resolved colours, in series order (`#rrggbb`). */
	colors: () => readonly string[];
}

/** The colourful palette (style 2) with every colour transformed. */
function recolour(transform: (color: string) => string): string[] {
	return getChartStylePalette(2).map(transform);
}

export const CHART_STYLE_PRESETS: readonly ChartStylePreset[] = [
	{ id: 'colorful', label: 'Colorful', styleId: 2, colors: () => getChartStylePalette(2) },
	{ id: 'monochrome', label: 'Monochrome', styleId: 3, colors: () => getChartStylePalette(3) },
	{
		id: 'colorfulLight',
		label: 'Colorful Light',
		styleId: 18,
		colors: () => recolour((c) => tint(c, 0.3)),
	},
	{
		id: 'colorfulDark',
		label: 'Colorful Dark',
		styleId: 26,
		colors: () => recolour((c) => shade(c, 0.3)),
	},
	{
		id: 'mutedDark',
		label: 'Muted Dark',
		styleId: 34,
		colors: () => recolour((c) => shade(tint(c, 0.3), 0.4)),
	},
	{ id: 'pastel', label: 'Pastel', styleId: 42, colors: () => recolour((c) => tint(c, 0.55)) },
];

/** The preset `id`, or `undefined`. */
export function findChartStylePreset(id: string): ChartStylePreset | undefined {
	return CHART_STYLE_PRESETS.find((preset) => preset.id === id);
}

/** Case-insensitive equality of two colour lists. */
export function sameChartColors(
	a: readonly (string | undefined)[],
	b: readonly (string | undefined)[] | undefined,
): boolean {
	return (
		!!b &&
		a.length === b.length &&
		a.every((c, i) => !!c && c.toLowerCase() === b[i]?.toLowerCase())
	);
}

/**
 * The preset whose colours `seriesColors` show, cycling the preset palette per series (a chart
 * whose series colours were pinned from a preset); `undefined` when none matches.
 */
export function chartStylePresetOfColors(
	seriesColors: readonly (string | undefined)[],
): ChartStylePreset | undefined {
	if (!seriesColors.length) return undefined;
	return CHART_STYLE_PRESETS.find((preset) => {
		const colors = preset.colors();
		return sameChartColors(
			seriesColors.map((_, i) => colors[i % colors.length]),
			seriesColors,
		);
	});
}
