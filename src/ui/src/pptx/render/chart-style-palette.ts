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
 * The gallery itself lives in `ooxml-core/chart` (`style-palette`).
 *
 * @module chart-style-palette
 */
import { DEFAULT_CHART_PALETTE, getChartStylePalette, shade, tint } from 'ooxml-core/chart';
import type { PptxChartData } from 'ooxml-core/pptx';

export { DEFAULT_CHART_PALETTE, getChartStylePalette, shade, tint };

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
