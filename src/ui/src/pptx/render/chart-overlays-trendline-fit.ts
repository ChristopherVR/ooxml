/**
 * chart-overlays-trendline-fit.ts: per-series trendline point computation
 * (linear / exponential / logarithmic / power / polynomial / moving-average).
 * The fit itself is `fitChartTrendline` in `ooxml-core/chart`, in data space;
 * this module chooses the x values and maps the fitted samples to pixels.
 *
 * Excel fits a trendline against x = 1, 2, ..., n on a category chart (bar /
 * line / area): the category POSITION, never a 0-based array index. A scatter
 * chart has no categories at all; Excel fits against the series' own real
 * `c:xVal` data instead. Collapsing both cases to a bare 0-based index (the
 * previous behaviour here) shifted every category-chart fit by one unit and
 * made a scatter trendline ignore its own X axis entirely, fitting against
 * point ORDER instead of the X values a reader actually sees on the chart.
 *
 * Ported / adapted from:
 *   packages/react/src/viewer/utils/chart-trendlines.tsx (regression engine)
 *   packages/shared/src/render/chart-trendlines.ts (shared port)
 *
 * @module chart-overlays-trendline-fit
 */

import { fitChartTrendline } from 'ooxml-core/chart';
import type { PptxChartSeries, PptxChartTrendline } from 'ooxml-core/pptx';

import type { PlotLayout, ValueRange } from './chart-view-model';
import { valueToY } from './chart-view-model';

/**
 * Map a (possibly fractional / extrapolated) category index to an X pixel.
 * `mode === 'bar'` centres on category slots; `'line'` anchors at data points.
 * Mirrors `xToPixel` from chart-overlay-utils.ts (React) and the shared port.
 */
export function xToPixel(
	xVal: number,
	catCount: number,
	layout: PlotLayout,
	mode: 'line' | 'bar',
): number {
	if (mode === 'bar') {
		const slotWidth = layout.plotWidth / Math.max(catCount, 1);
		return layout.plotLeft + slotWidth * xVal + slotWidth / 2;
	}
	const maxIdx = Math.max(catCount - 1, 1);
	return layout.plotLeft + (xVal / maxIdx) * layout.plotWidth;
}

export interface ComputedTrend {
	points: { x: number; y: number }[];
	equation: string;
	rSquared: number;
}

/** A scatter chart's shared X-axis domain, matching `computeScatterDots`' own normalisation. */
export interface TrendlineXDomain {
	min: number;
	span: number;
}

/**
 * Compute the polyline points (and equation / R-squared) for one trendline
 * over a series. Returns empty points when the type is unsupported or data
 * is too sparse.
 *
 * @param xDomain  A scatter chart's X-axis domain (from `computeScatterXDomain`),
 *   present only when `series.xValues` should be used both for the fit AND for
 *   placing the drawn curve. Absent (category charts) fits against x = 1..n and
 *   places the curve via the existing 0-based `xToPixel` category convention.
 */
export function computeTrendlinePoints(
	trendline: PptxChartTrendline,
	series: PptxChartSeries,
	catCount: number,
	layout: PlotLayout,
	range: ValueRange,
	mode: 'line' | 'bar',
	xDomain?: TrendlineXDomain,
): ComputedTrend {
	const yVals = series.values;
	const n = yVals.length;
	if (n < 2) {
		return { points: [], equation: '', rSquared: 0 };
	}

	const useRealX = xDomain !== undefined && (series.xValues?.length ?? 0) >= n;
	// Fit-space X: the series' own real data for a scatter chart, else Excel's
	// 1-based category position (never a bare 0-based index).
	const xVals = useRealX ? series.xValues!.slice(0, n) : yVals.map((_v, i) => i + 1);

	// A fit-space X maps to a pixel either through the scatter domain's own
	// normalisation (matching `computeScatterDots`, so the curve lines up with
	// the actual plotted points) or, for a category chart, through the
	// existing 0-based `xToPixel` convention (fitX is 1-based, so shift back).
	const toPixelX = (fitX: number): number =>
		useRealX
			? layout.plotLeft +
				(xDomain!.span > 0 ? (fitX - xDomain!.min) / xDomain!.span : 0) * layout.plotWidth
			: xToPixel(fitX - 1, catCount, layout, mode);

	const fit = fitChartTrendline(
		{
			kind: trendline.trendlineType,
			order: trendline.order,
			period: trendline.period,
			forward: trendline.forward,
			backward: trendline.backward,
			intercept: trendline.intercept,
		},
		xVals,
		yVals,
	);
	const points = fit.points.map((point) => ({
		x: toPixelX(point.x),
		y: valueToY(point.y, range, layout.plotTop, layout.plotBottom),
	}));
	return { points, equation: fit.equation, rSquared: fit.rSquared };
}
