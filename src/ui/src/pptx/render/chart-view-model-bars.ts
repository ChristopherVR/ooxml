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
		singleBarWidth = (barGroupWidth * 0.7) / seriesCount,
		groupOffset = (barGroupWidth - singleBarWidth * seriesCount) / 2;

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
 * one; without it the bar keeps its old `fallbackFraction` of the slot.
 */
export function stackedBarWidth(
	slotWidth: number,
	gapWidth: number | undefined,
	fallbackFraction: number,
): number {
	return gapWidth === undefined
		? slotWidth * fallbackFraction
		: clusteredBarGeometry(slotWidth, 1, { barGapWidth: gapWidth }).singleBarWidth;
}

export function computeStackedBarRects(
	series: ReadonlyArray<PptxChartSeries>,
	catCount: number,
	layout: PlotLayout,
	range: ValueRange,
	colorPalette: readonly string[] | undefined,
	gapWidth?: number,
): BarRect[] {
	const rects: BarRect[] = [],
		barW = stackedBarWidth(layout.plotWidth / Math.max(catCount, 1), gapWidth, 0.7),
		barOffset = (layout.plotWidth / Math.max(catCount, 1) - barW) / 2,
		zeroY = valueToY(0, range, layout.plotTop, layout.plotBottom);

	for (let ci = 0; ci < catCount; ci++) {
		let posTop = zeroY,
			negBottom = zeroY;

		for (let si = 0; si < series.length; si++) {
			const val = series[si].values[ci] ?? 0;
			if (val === 0) {
				continue;
			}
			const x = layout.plotLeft + (layout.plotWidth / Math.max(catCount, 1)) * ci + barOffset,
				h = Math.max(
					Math.abs(
						valueToY(val, range, layout.plotTop, layout.plotBottom) -
							valueToY(0, range, layout.plotTop, layout.plotBottom),
					),
					1,
				);
			if (val > 0) {
				const y = posTop - h;
				rects.push({
					x,
					y,
					w: barW,
					h,
					fill: seriesColor(series[si], si, colorPalette),
					seriesIndex: si,
					pointIndex: ci,
				});
				posTop = y;
			} else {
				const y = negBottom;
				rects.push({
					x,
					y,
					w: barW,
					h,
					fill: seriesColor(series[si], si, colorPalette),
					seriesIndex: si,
					pointIndex: ci,
				});
				negBottom = y + h;
			}
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
