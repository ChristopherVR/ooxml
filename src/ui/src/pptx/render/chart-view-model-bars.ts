/**
 * chart-view-model-bars.ts: bar / column rectangle and line-point geometry of
 * the chart engine. Split out of `chart-view-model.ts`, which re-exports
 * everything here.
 *
 * @module chart-view-model-bars
 */
/* eslint-disable one-var -- this module predates the rule and combining every
   sibling `const`/`let` in a function into one comma-list (oxlint's own
   `--fix` cannot do this safely once a non-declaration statement sits between
   them) would churn geometry code far beyond this change's scope. */

import type { PptxChartSeries } from 'ooxml-core/pptx';

import { clusteredBarGeometry } from './chart-bar-cluster-geometry';
import { seriesColor, valueToY } from './chart-view-model-scale';
import type { ValueRange } from './chart-view-model-scale';
import type { PlotLayout } from './chart-view-model-types';

// ─────────────────────────────────────────────────────────────────────────────
// Bar / column
// ─────────────────────────────────────────────────────────────────────────────

export interface BarRect {
	x: number;
	y: number;
	w: number;
	h: number;
	fill: string;
	/** Source series index, carried so plot builders can tag interactive parts. */
	seriesIndex?: number;
	/** Source category index, carried so plot builders can tag interactive parts. */
	pointIndex?: number;
}

export function computeBarRects(
	series: ReadonlyArray<PptxChartSeries>,
	catCount: number,
	layout: PlotLayout,
	range: ValueRange,
	colorPalette: readonly string[] | undefined,
): BarRect[] {
	const rects: BarRect[] = [],
		seriesCount = Math.max(series.length, 1),
		barGroupWidth = layout.plotWidth / Math.max(catCount, 1),
		{ singleBarWidth, clusterWidth } = clusteredBarGeometry(barGroupWidth, seriesCount, {}),
		groupOffset = (barGroupWidth - clusterWidth) / 2;

	for (let ci = 0; ci < catCount; ci++) {
		for (let si = 0; si < series.length; si++) {
			const val = series[si].values[ci] ?? 0,
				x = layout.plotLeft + barGroupWidth * ci + groupOffset + singleBarWidth * si,
				zeroY = valueToY(0, range, layout.plotTop, layout.plotBottom),
				valY = valueToY(val, range, layout.plotTop, layout.plotBottom),
				y = Math.min(zeroY, valY),
				h = Math.max(Math.abs(zeroY - valY), 1);
			rects.push({
				x,
				y,
				w: singleBarWidth,
				h,
				fill: seriesColor(series[si], si, colorPalette),
			});
		}
	}
	return rects;
}

/**
 * Width of a stacked bar in one category slot. All the series of a stacked
 * chart share one bar, so `c:gapWidth` sizes it the way it sizes a cluster of
 * one. An absent `c:gapWidth` takes the ECMA-376 default of 150%
 * (`DEFAULT_BAR_GAP_WIDTH`), so the bar is 40% of the slot.
 */
export function stackedBarWidth(slotWidth: number, gapWidth: number | undefined): number {
	return clusteredBarGeometry(slotWidth, 1, gapWidth === undefined ? {} : { barGapWidth: gapWidth })
		.singleBarWidth;
}

/**
 * Clip one stacked segment, running from `base` to `base + value` on the value
 * axis, to the axis range. An explicit `c:min` / `c:max` cuts the segment at
 * the plot edge, as PowerPoint draws it, and a segment wholly outside the
 * range is dropped (`undefined`). Every stacked bar path goes through here.
 */
export function clipStackedSegment(
	base: number,
	value: number,
	range: ValueRange,
): { low: number; high: number } | undefined {
	const low = Math.max(Math.min(base, base + value), range.min),
		high = Math.min(Math.max(base, base + value), range.max);
	return high > low ? { low, high } : undefined;
}

/**
 * Pixel extent of a clipped stacked segment between `baseCoord` (where it
 * sits on the stack) and `tipCoord`. A segment thinner than `minSize` grows
 * away from its base, so it never covers the segment it sits on.
 */
export function stackedSegmentExtent(
	baseCoord: number,
	tipCoord: number,
	minSize: number,
): { start: number; size: number } {
	const size = Math.max(Math.abs(tipCoord - baseCoord), minSize);
	return { start: tipCoord < baseCoord ? baseCoord - size : baseCoord, size };
}

/**
 * One rect per non-zero segment of a stacked column chart: each series sits on
 * the running total of the same-sign series below it, and the segment is
 * clipped to the value axis by {@link clipStackedSegment}.
 */
export function computeStackedBarRects(
	series: ReadonlyArray<PptxChartSeries>,
	catCount: number,
	layout: PlotLayout,
	range: ValueRange,
	colorPalette: readonly string[] | undefined,
	gapWidth?: number,
): BarRect[] {
	const rects: BarRect[] = [],
		slot = layout.plotWidth / Math.max(catCount, 1),
		barW = stackedBarWidth(slot, gapWidth),
		barOffset = (slot - barW) / 2;

	for (let ci = 0; ci < catCount; ci++) {
		let posTotal = 0,
			negTotal = 0;

		for (let si = 0; si < series.length; si++) {
			const val = series[si].values[ci] ?? 0;
			if (val === 0) {
				continue;
			}
			const base = val > 0 ? posTotal : negTotal;
			if (val > 0) {
				posTotal += val;
			} else {
				negTotal += val;
			}
			const segment = clipStackedSegment(base, val, range);
			if (!segment) {
				continue;
			}
			const baseY = valueToY(
					val > 0 ? segment.low : segment.high,
					range,
					layout.plotTop,
					layout.plotBottom,
				),
				tipY = valueToY(
					val > 0 ? segment.high : segment.low,
					range,
					layout.plotTop,
					layout.plotBottom,
				),
				{ start, size } = stackedSegmentExtent(baseY, tipY, 1);
			rects.push({
				x: layout.plotLeft + slot * ci + barOffset,
				y: start,
				w: barW,
				h: size,
				fill: seriesColor(series[si], si, colorPalette),
				seriesIndex: si,
				pointIndex: ci,
			});
		}
	}
	return rects;
}

// ─────────────────────────────────────────────────────────────────────────────
// Line / area
// ─────────────────────────────────────────────────────────────────────────────

export interface LinePoint {
	x: number;
	y: number;
}

export function computeLinePoints(
	values: ReadonlyArray<number>,
	catCount: number,
	layout: PlotLayout,
	range: ValueRange,
): LinePoint[] {
	const n = Math.max(catCount, 2);
	return values.map((val, i) => {
		const nx = n > 1 ? i / (n - 1) : 0,
			x = layout.plotLeft + layout.plotWidth * nx,
			y = valueToY(val, range, layout.plotTop, layout.plotBottom);
		return { x, y };
	});
}

export function linePointsToSvgString(points: ReadonlyArray<LinePoint>): string {
	return points.map((p) => `${p.x.toFixed(2)},${p.y.toFixed(2)}`).join(' ');
}
