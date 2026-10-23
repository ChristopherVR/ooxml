/**
 * chart-quick-action-styles.ts: PowerPoint's "Chart Styles" quick-action
 * gallery (the paintbrush icon shown outside a selected chart), a one-click
 * recolour of the whole chart.
 *
 * No preset/gallery model existed anywhere in this codebase before this
 * module: `chart-style-defaults.ts` only resolves font/colour FALLBACKS from
 * a chart's own parsed `chartStyleDefinition` part, and `chart-helpers.ts`'s
 * `getChartStylePalette(styleId)` (1-48) generates the 2D chart engine's
 * built-in style palettes, but nothing curated a small user-facing subset of
 * them or applied one to a chart. This module is deliberately modest (six
 * presets) rather than PowerPoint's full 48+ gallery, per this feature's
 * scope: two presets are Office style palettes from `getChartStylePalette`
 * and four are recolours of the colourful one, not a new colour model. The
 * preset table itself is shared by every product's Chart Design gallery and
 * lives in `ooxml-core/chart` (`style-presets`); this module keeps the
 * `PptxChartData` mapping.
 *
 * The 2D SVG chart engine (`chart-view-model-scale.ts`'s `seriesColor`/
 * `paletteColor`) reads `PptxChartData.colorPalette` directly and does NOT
 * consult `style.styleId` (that field only feeds the separate 3-D chart
 * engine's palette resolution, see `chart-view-model.ts`'s header comment).
 * So {@link applyChartStylePreset} writes the resolved palette straight onto
 * `colorPalette` (which the doc comment on that field already says takes
 * priority over the styleId-derived one), and also mirrors `style.styleId`
 * so the 3-D engine and the "which preset is applied" check stay consistent.
 *
 * `seriesColor()` (`chart-view-model-scale.ts`) and `resolveDataPointFill()`
 * (`chart-datapoint-style.ts`) both prefer an explicit per-series/per-point
 * colour over the palette, and virtually every chart authored by PowerPoint,
 * Excel or WPS sets one (`c:ser/c:spPr` and/or `c:marker/c:spPr`), because
 * that is how "no palette, pick your own colours" charts are normally
 * authored. Writing only `colorPalette` therefore left the whole feature a
 * no-op for real-world decks: the freshly-written palette was always shadowed
 * by the pre-existing explicit colours, and clicking a preset visibly changed
 * nothing (reported after this shipped). A "Chart Styles" click is meant to
 * be a one-click WHOLE-CHART recolour, matching PowerPoint's own behaviour
 * for its Style/Colour gallery, so {@link applyChartStylePreset} also clears
 * every series' `color` and its marker's `spPr.fillColor`, plus every
 * per-point `c:dPt` fill override (`dataPoints[].spPr.fillColor` and
 * `dataPoints[].marker.spPr.fillColor`), so nothing is left shadowing the new
 * palette. Only fill colours are cleared: marker symbol/size, stroke/border
 * colours (e.g. a marker's white ring), dash styles, and everything else
 * about the series/point are left untouched.
 *
 * @module render/chart-quick-action-styles
 */
import type { PptxChartData, PptxChartDataPoint, PptxChartSeries } from 'ooxml-core/pptx';

import { CHART_STYLE_PRESETS, findChartStylePreset, sameChartColors } from 'ooxml-core/chart';

/** One built-in "Chart Styles" gallery entry. */
export interface ChartStylePresetDescriptor {
	id: string;
	labelKey: string;
	/** Resolved swatch colours (for the gallery thumbnail), in series order. */
	colors: readonly string[];
	/** Whether this preset's palette matches the chart's current `colorPalette`. */
	applied: boolean;
}

/** The i18n key of a preset's quick-action label (`colorfulLight` is `pptx.chart.styleColorfulLight`). */
const labelKeyOf = (id: string): string => `pptx.chart.style${id[0]!.toUpperCase()}${id.slice(1)}`;

/** The resolved colours of a "Chart Styles" preset, or `undefined` for an unknown id. */
export function chartStylePresetPalette(id: string): readonly string[] | undefined {
	return findChartStylePreset(id)?.colors();
}

/**
 * Build the "Chart Styles" gallery descriptor for the currently-selected
 * chart: every preset's resolved colours plus whether it is the one
 * currently applied (its resolved palette matches `chartData.colorPalette`).
 */
export function buildChartStylePresets(
	chartData: Pick<PptxChartData, 'colorPalette'> & Partial<Pick<PptxChartData, 'series'>>,
): ChartStylePresetDescriptor[] {
	return CHART_STYLE_PRESETS.map((preset) => {
		const colors = preset.colors();
		return {
			id: preset.id,
			labelKey: labelKeyOf(preset.id),
			colors,
			applied:
				sameChartColors(colors, chartData.colorPalette) ||
				(!chartData.colorPalette &&
					Boolean(chartData.series?.length) &&
					chartData.series!.every(
						(series, index) =>
							series.color?.toLowerCase() === colors[index % colors.length]?.toLowerCase(),
					)),
		};
	});
}

/** Drop a data point's own fill override, keeping every other field. */
function clearDataPointFill(point: PptxChartDataPoint): PptxChartDataPoint {
	const { spPr, marker, ...rest } = point;
	const nextSpPr = spPr && { ...spPr, fillColor: undefined };
	const nextMarker = marker?.spPr && {
		...marker,
		spPr: { ...marker.spPr, fillColor: undefined },
	};
	return { ...rest, ...(spPr && { spPr: nextSpPr }), ...(nextMarker && { marker: nextMarker }) };
}

/** Drop a series' own colour overrides (and its points'), keeping shape/symbol/size. */
export function clearSeriesColor(series: PptxChartSeries): PptxChartSeries {
	const { color: _color, marker, dataPoints, ...rest } = series;
	const nextMarker = marker?.spPr && { ...marker, spPr: { ...marker.spPr, fillColor: undefined } };
	return {
		...rest,
		...(marker && { marker: nextMarker ?? marker }),
		...(dataPoints && { dataPoints: dataPoints.map(clearDataPointFill) }),
	};
}

/**
 * Apply a "Chart Styles" preset by id: resolves its palette via
 * its preset entry, writes it onto `colorPalette` (also mirroring
 * `style.styleId`), and clears every series'/data-point's own explicit fill
 * so the new palette is not immediately shadowed by pre-existing colours
 * (see this module's header). Returns `null` for an unknown preset id so a
 * caller can no-op instead of silently clearing the chart's palette.
 */
export function applyChartStylePreset(
	chartData: PptxChartData,
	presetId: string,
): PptxChartData | null {
	const preset = findChartStylePreset(presetId);
	if (!preset) {
		return null;
	}
	const colors = [...preset.colors()];
	return {
		...chartData,
		colorPalette: colors,
		style: { ...chartData.style, styleId: preset.styleId },
		series: chartData.series.map(clearSeriesColor),
	};
}
