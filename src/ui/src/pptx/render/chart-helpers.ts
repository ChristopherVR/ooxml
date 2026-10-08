import type {
	PptxChartData,
	PptxChartSeries,
	PptxChartStyle,
	PptxChartAxisFormatting,
} from 'ooxml-core/pptx';

import { axisTargetIntervals, niceValueAxisBounds } from './chart-axis-nice';
import { chartFontPx } from './chart-font';
import { reserveLegendSpace } from './chart-legend-placement';
import { formatChartNumber } from './chart-number-format';
import { getChartStylePalette } from './chart-style-palette';
import { chartTitleBand } from './chart-title-band';

/**
 * Framework-agnostic chart helpers, a focused Vue port of the React package's
 * `viewer/utils/chart-helpers.ts`, `chart-layout.ts`, and
 * `chart-style-palettes.ts`.
 *
 * This module only covers the common-type renderers (bar / column / stacked /
 * line / area / pie / doughnut) and the shared chrome. Advanced axis/overlay
 * features (log axes, secondary axes, display units, trendlines, error bars,
 * data tables) live in `chart-axis.ts` / `chart-axis-render.ts` /
 * `chart-cartesian.ts` and are consumed via `chart/ChartViewModelSvg.vue`.
 *
 * These small pure helpers are an extraction candidate: long-term they (and
 * their React counterparts) should live in a shared, framework-agnostic
 * package so all three UI bindings reuse one implementation.
 */

// ── Style palettes ───────────────────────────────────────────────

export { DEFAULT_CHART_PALETTE, getChartStylePalette } from './chart-style-palette';

/**
 * Resolve a series colour: explicit colour → marker fill → parsed palette →
 * style palette. The marker fallback covers scatter series authored with
 * `c:ser/c:spPr/a:ln/a:noFill` plus a coloured `c:marker/c:spPr`: the points
 * paint the marker fill directly, so the legend swatch must match it rather
 * than fall back to the palette.
 */
export function seriesColor(
	series: PptxChartSeries,
	index: number,
	styleId?: number,
	colorPalette?: string[],
): string {
	const explicit = series.color ?? series.marker?.spPr?.fillColor;
	if (explicit) {
		return explicit;
	}
	if (colorPalette && colorPalette.length > 0) {
		return colorPalette[index % colorPalette.length];
	}
	const palette = getChartStylePalette(styleId);
	return palette[index % palette.length];
}

/** Palette colour for an index with no series object (e.g. per-slice pie colouring). */
export function paletteColor(index: number, styleId?: number, colorPalette?: string[]): string {
	if (colorPalette && colorPalette.length > 0) {
		return colorPalette[index % colorPalette.length];
	}
	const palette = getChartStylePalette(styleId);
	return palette[index % palette.length];
}

// ── Value range ──────────────────────────────────────────────────

export interface ValueRange {
	min: number;
	max: number;
	span: number;
	/** When true, the range represents log-scaled values (see `chart-axis.ts`). */
	logScale?: boolean;
	/** Logarithmic base (e.g. 10, 2, Math.E). Only meaningful when logScale is true. */
	logBase?: number;
	/** Whether values increase from top to bottom (`c:orientation="maxMin"`). */
	reverseOrder?: boolean;
	/**
	 * Step between major gridlines when the bounds came from the automatic
	 * scale, which always spans a whole number of them. Tick generators use it
	 * instead of dividing the span evenly, which is what keeps the labels on
	 * round numbers. Absent when an explicit `c:min`/`c:max` overrode the
	 * automatic bounds, since the unit no longer divides them.
	 */
	majorUnit?: number;
}

/**
 * Automatic Y-axis range, on PowerPoint's terms: zero-anchored where that reads
 * sensibly, padded, and rounded out to whole major units. See
 * `chart-axis-nice.ts` for the rules. Running the axis to the raw data maximum
 * instead put the top gridline on whatever the tallest bar happened to be.
 *
 * @param plotHeightPx The chart's plot-area height (same "px" unit
 *   `computePlotLayout` returns), when the caller has it. PowerPoint's real
 *   gridline count depends on how much vertical room the axis has to label
 *   them in; omitting this falls back to a constant tuned for a short chart
 *   (see `axisTargetIntervals`'s doc comment).
 */
export function computeValueRange(
	series: ReadonlyArray<PptxChartSeries>,
	plotHeightPx?: number,
): ValueRange {
	const allValues = series.flatMap((s) => s.values);
	if (allValues.length === 0) {
		return { min: 0, max: 1, span: 1 };
	}
	const { min, max, majorUnit } = niceValueAxisBounds(
		Math.min(...allValues),
		Math.max(...allValues),
		plotHeightPx !== undefined ? axisTargetIntervals(plotHeightPx) : undefined,
	);
	return { min, max, span: Math.max(max - min, Number.EPSILON), majorUnit };
}

/**
 * Map a data value to a Y pixel coordinate (top = max, bottom = min).
 * Routes through logarithmic scaling when `range.logScale` is set (the log
 * helpers live in `chart-axis.ts`; the branch is inlined here to avoid a
 * circular import).
 */
export function valueToY(val: number, range: ValueRange, topY: number, bottomY: number): number {
	const usable = bottomY - topY;
	let ratio: number;
	if (range.logScale && range.logBase) {
		const base = range.logBase;
		const clampedVal = Math.max(val, range.min);
		const logVal = Math.log(clampedVal) / Math.log(base);
		const logMin = Math.log(range.min) / Math.log(base);
		ratio = (logVal - logMin) / range.span;
	} else {
		ratio = (val - range.min) / range.span;
	}
	return range.reverseOrder ? topY + ratio * usable : bottomY - ratio * usable;
}

/**
 * Compact axis-value formatting (1.2K / 3.4M / integer / one-decimal).
 *
 * `formatCode` is the chart's own `c:numFmt/@formatCode`. When the source
 * declares one - `0%`, `#,##0`, `$#,##0.00` - it WINS: the cached values behind
 * a percentage chart are fractions, and the compact fallback renders them as
 * `0.5` where PowerPoint shows `50%`. Codes outside the supported subset fall
 * through to the compact form, which is what every caller did before.
 */
export function formatAxisValue(val: number, formatCode?: string): string {
	const formatted = formatChartNumber(val, formatCode);
	if (formatted !== undefined) {
		return formatted;
	}
	if (Math.abs(val) >= 1_000_000) {
		return `${(val / 1_000_000).toFixed(1)}M`;
	}
	if (Math.abs(val) >= 1_000) {
		return `${(val / 1_000).toFixed(1)}K`;
	}
	if (Number.isInteger(val)) {
		return String(val);
	}
	return val.toFixed(1);
}

// ── Layout ───────────────────────────────────────────────────────

export interface PlotLayout {
	plotLeft: number;
	plotTop: number;
	plotRight: number;
	plotBottom: number;
	plotWidth: number;
	plotHeight: number;
	svgWidth: number;
	svgHeight: number;
}

/**
 * Reserved-space options for `computeLayout` (secondary axes + data table).
 * Structurally identical to `LayoutOptions` in `chart-axis.ts`; declared here
 * as a plain shape to avoid a circular import (chart-axis depends on this file).
 */
export interface ComputeLayoutOptions {
	hasSecondaryValueAxis?: boolean;
	hasSecondaryCategoryAxis?: boolean;
	hasDataTable?: boolean;
	dataTableRowCount?: number;
}

/**
 * Compute the plot rectangle within the SVG, reserving room for axes, title,
 * legend, optional secondary axes, and an optional data table.
 */
export function computeLayout(
	elementWidth: number,
	elementHeight: number,
	style: PptxChartStyle | undefined,
	hasAxes: boolean,
	legendPos: string,
	options?: ComputeLayoutOptions,
): PlotLayout {
	// Match the element frame exactly; bindings stretch the viewBox with
	// preserveAspectRatio "none", so a minimum would scale non-uniformly
	// (see computePlotLayout in chart-view-model.ts).
	const svgWidth = Math.max(1, elementWidth);
	const svgHeight = Math.max(1, elementHeight);
	let plotLeft = hasAxes ? 48 : 8;
	let plotTop = 8;
	let plotRight = svgWidth - 8;
	let plotBottom = svgHeight - (hasAxes ? 24 : 8);

	if (style?.hasTitle) {
		plotTop += chartTitleBand(
			style.titleFontSize !== undefined ? chartFontPx(style.titleFontSize) : 12,
		).bandHeight;
	}
	if (style?.hasLegend) {
		// `tr` (top-right corner) overlays the plot per PowerPoint's own
		// quick-layout behaviour: no band is reserved for it.
		({ plotLeft, plotTop, plotRight, plotBottom } = reserveLegendSpace(legendPos, {
			plotLeft,
			plotTop,
			plotRight,
			plotBottom,
		}));
	}

	// Reserve space for a secondary value axis on the right.
	if (options?.hasSecondaryValueAxis) {
		plotRight -= 40;
	}
	// Reserve space for a secondary category axis on the top.
	if (options?.hasSecondaryCategoryAxis) {
		plotTop += 16;
	}
	// Reserve space for a data table below the chart.
	if (options?.hasDataTable) {
		const rowCount = options.dataTableRowCount ?? 1;
		plotBottom -= 14 + rowCount * 14;
	}

	const pw = Math.max(plotRight - plotLeft, 1);
	const ph = Math.max(plotBottom - plotTop, 1);
	return {
		plotLeft,
		plotTop,
		plotRight: plotLeft + pw,
		plotBottom: plotTop + ph,
		plotWidth: pw,
		plotHeight: ph,
		svgWidth,
		svgHeight,
	};
}

// ── Derived helpers ──────────────────────────────────────────────

/** Find the primary value-axis formatting (non-right, or first valAx). */
export function getPrimaryValueAxis(
	axes: PptxChartAxisFormatting[] | undefined,
): PptxChartAxisFormatting | undefined {
	if (!axes) {
		return undefined;
	}
	return (
		axes.find((a) => a.axisType === 'valAx' && a.axPos !== 'r') ??
		axes.find((a) => a.axisType === 'valAx')
	);
}

/**
 * Resolve category labels for a chart: use the declared categories, otherwise
 * synthesise 1-based labels for the longest series.
 */
export function resolveCategoryLabels(chartData: PptxChartData): string[] {
	if (chartData.categories.length > 0) {
		return chartData.categories;
	}
	const longest = chartData.series.reduce((max, s) => Math.max(max, s.values.length), 0);
	return Array.from({ length: longest }, (_, i) => String(i + 1));
}
