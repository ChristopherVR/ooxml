/**
 * chart-data-helpers.ts: Pure immutable ELEMENT-level wrappers for chart data
 * editing (`ChartPptxElement` in, `ChartPptxElement` out).
 *
 * The chart data sheet's row/column edits go through the shared
 * `pptx-ui-chart-data-sheet` element and core's `chartDataApplySheetEdits`;
 * what stays here is the per-series colour, `patchChartStyle` / `patchChartData`
 * and the advanced formatting wrappers, which have no shared element yet.
 *
 * All functions are immutable (return new objects, leave inputs unchanged)
 * and framework-agnostic.
 *
 * Ported from the React inspector:
 *   packages/react/src/viewer/components/inspector/ChartDataPanel.tsx
 *
 * @module angular-viewer/chart-data-helpers
 */

import { patchChartData as sharedPatchChartData } from 'ooxml-ui/pptx';
import {
	chartDataChangeType,
	chartDataUpdatePoint,
	setChartAxisGridlineStyle,
	setChartAxisLogScale,
	setChartAxisTitleStyle,
	setChartDataPointExplosion,
	setChartDataPointFill,
	setChartDataPointMarker,
	setChartSeriesChartType,
	setChartSeriesMarker,
} from 'pptx-viewer-core';
import type {
	ChartAxisTitleStyleEdit,
	ChartGridlineStyleEdit,
	ChartPptxElement,
	PptxChartAxisType,
	PptxChartData,
	PptxChartMarkerSymbol,
	PptxChartSeries,
	PptxChartStyle,
	PptxChartType,
} from 'pptx-viewer-core';

// Re-export core primitives so callers can import everything from one place.
export { chartDataChangeType, chartDataUpdatePoint };

// ---------------------------------------------------------------------------
// Advanced formatting wrappers (log scale, title/gridline style, markers,
// combo per-series type, per-point fill/explosion).
//
// Each clones the element, runs the in-place core op, and returns a new
// `ChartPptxElement` so the Angular editor's immutable contract is preserved.
// These make the new chart-editing ops consumable from the Angular binding even
// though the dedicated inspector controls for them are not yet ported.
// ---------------------------------------------------------------------------

/** Apply an in-place core chart op to a deep clone of `element`. */
function withClonedChart(
	element: ChartPptxElement,
	mutate: (clone: ChartPptxElement) => void,
): ChartPptxElement {
	if (!element.chartData) {
		return element;
	}
	const clone: ChartPptxElement = {
		...element,
		chartData: structuredClone(element.chartData),
	};
	mutate(clone);
	return clone;
}

/** Enable/disable logarithmic scaling on an axis. */
export function setAxisLogScale(
	element: ChartPptxElement,
	axisType: PptxChartAxisType,
	opts: { enabled: boolean; base?: number },
): ChartPptxElement {
	return withClonedChart(element, (el) => setChartAxisLogScale(el, axisType, opts));
}

/** Edit an axis title's font styling. */
export function setAxisTitleStyle(
	element: ChartPptxElement,
	axisType: PptxChartAxisType,
	edit: ChartAxisTitleStyleEdit,
): ChartPptxElement {
	return withClonedChart(element, (el) => setChartAxisTitleStyle(el, axisType, edit));
}

/** Edit major/minor gridline line styling for an axis. */
export function setGridlineStyle(
	element: ChartPptxElement,
	axisType: PptxChartAxisType,
	which: 'major' | 'minor',
	edit: ChartGridlineStyleEdit,
): ChartPptxElement {
	return withClonedChart(element, (el) => setChartAxisGridlineStyle(el, axisType, which, edit));
}

/** Set or clear a series marker. */
export function setSeriesMarker(
	element: ChartPptxElement,
	seriesIndex: number,
	marker: { symbol?: PptxChartMarkerSymbol; size?: number; fillColor?: string } | null,
): ChartPptxElement {
	return withClonedChart(element, (el) => setChartSeriesMarker(el, seriesIndex, marker));
}

/** Set or clear a per-series chart type (combo charts). */
export function setSeriesChartType(
	element: ChartPptxElement,
	seriesIndex: number,
	seriesType: PptxChartType | null,
): ChartPptxElement {
	return withClonedChart(element, (el) => setChartSeriesChartType(el, seriesIndex, seriesType));
}

/** Set or clear a per-data-point fill colour. */
export function setDataPointFill(
	element: ChartPptxElement,
	seriesIndex: number,
	pointIndex: number,
	color: string | null,
): ChartPptxElement {
	return withClonedChart(element, (el) =>
		setChartDataPointFill(el, seriesIndex, pointIndex, color),
	);
}

/**
 * Set or clear a per-data-point marker override (`c:dPt/c:marker`), which
 * replaces the series marker for that one point.
 */
export function setDataPointMarker(
	element: ChartPptxElement,
	seriesIndex: number,
	pointIndex: number,
	marker: { symbol?: PptxChartMarkerSymbol; size?: number; fillColor?: string } | null,
): ChartPptxElement {
	return withClonedChart(element, (el) =>
		setChartDataPointMarker(el, seriesIndex, pointIndex, marker),
	);
}

/** Set or clear a per-data-point pie/doughnut slice explosion. */
export function setDataPointExplosion(
	element: ChartPptxElement,
	seriesIndex: number,
	pointIndex: number,
	explosion: number | null,
): ChartPptxElement {
	return withClonedChart(element, (el) =>
		setChartDataPointExplosion(el, seriesIndex, pointIndex, explosion),
	);
}

// ---------------------------------------------------------------------------
// setSeriesColor
// ---------------------------------------------------------------------------

/**
 * Set (or clear) the solid fill colour of a chart series, returning a new
 * `ChartPptxElement`. Pass a hex string (`#RRGGBB` or `RRGGBB`) to set the
 * colour, or `null` to clear it so the series falls back to its theme colour.
 *
 * Mirrors the headless `setChartSeriesColor` SDK op (which mutates in place)
 * with an immutable element-level wrapper for the inspector.
 *
 * @param element - The source chart element (not mutated).
 * @param seriesIndex - Zero-based index of the series to recolour.
 * @param color - Hex colour string, or `null` to clear.
 * @returns A new `ChartPptxElement`.
 *
 * @example
 * ```ts
 * const updated = setSeriesColor(el, 0, "#4472C4");
 * ```
 */
export function setSeriesColor(
	element: ChartPptxElement,
	seriesIndex: number,
	color: string | null,
): ChartPptxElement {
	const chartData = element.chartData;
	if (!chartData) {
		return element;
	}
	const normalized = color ? normalizeHex(color) : undefined;
	const series = chartData.series.map((s, i): PptxChartSeries =>
		i === seriesIndex ? { ...s, color: normalized } : s,
	);
	return { ...element, chartData: { ...chartData, series } };
}

/** Normalise a hex colour to a `#`-prefixed form, trimming whitespace. */
function normalizeHex(color: string): string {
	const trimmed = color.trim();
	return trimmed.startsWith('#') ? trimmed : `#${trimmed}`;
}

// ---------------------------------------------------------------------------
// patchChartStyle
// ---------------------------------------------------------------------------

/**
 * Merge a `Partial<PptxChartStyle>` patch into a `ChartPptxElement`,
 * returning a new element.
 *
 * @param element - The source chart element (not mutated).
 * @param patch - Style fields to merge.
 * @returns A new `ChartPptxElement`.
 */
export function patchChartStyle(
	element: ChartPptxElement,
	patch: Partial<PptxChartStyle>,
): ChartPptxElement {
	const chartData = element.chartData;
	if (!chartData) {
		return element;
	}
	return {
		...element,
		chartData: { ...chartData, style: { ...chartData.style, ...patch } },
	};
}

// ---------------------------------------------------------------------------
// patchChartData
// ---------------------------------------------------------------------------

/**
 * Merge a `Partial<PptxChartData>` patch into a `ChartPptxElement`.
 *
 * When the patch contains a `chartType` change, the smart
 * `chartDataChangeType` helper is used so grouping and category formats
 * are adapted automatically.
 *
 * @param element - The source chart element (not mutated).
 * @param patch - Chart data fields to merge.
 * @returns A new `ChartPptxElement`.
 */
export function patchChartData(
	element: ChartPptxElement,
	patch: Partial<PptxChartData>,
): ChartPptxElement {
	const chartData = element.chartData;
	if (!chartData) {
		return element;
	}
	return { ...element, chartData: sharedPatchChartData(chartData, patch) };
}
