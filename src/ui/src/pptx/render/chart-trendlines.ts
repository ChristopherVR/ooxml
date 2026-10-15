/**
 * chart-trendlines.ts: the earlier public trendline entry of `ooxml-ui/pptx`.
 *
 * `computeTrendlinePoints` fits one series' trendline against the 0-based
 * point index and returns its polyline in pixels plus a two-decimal equation.
 * The fit is `fitChartTrendline` from `ooxml-core/chart`; the renderer itself
 * uses `chart-overlays-trendline-fit.ts`, which fits against Excel's 1-based
 * category positions or a scatter series' own x values.
 */
import { fitChartTrendline, type ChartTrendlineFit } from 'ooxml-core/chart';
import type { PptxChartTrendline } from 'ooxml-core/pptx';

import type { PlotLayout, ValueRange } from './chart-helpers';

/** A single point on a computed trendline, in SVG pixel space. */
export interface TrendlinePoint {
	x: number;
	y: number;
}

/**
 * Map a (possibly fractional / extrapolated) category index to an X pixel.
 * `mode === 'bar'` centres on category slots; `'line'` anchors at points.
 */
function xToPixel(
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

interface ComputedTrend {
	points: TrendlinePoint[];
	equation: string;
	rSquared: number;
}

/** The two-decimal equation text this entry has always returned. */
function equationText(type: PptxChartTrendline['trendlineType'], fit: ChartTrendlineFit): string {
	const [first = 0, second = 0] = fit.coefficients;
	switch (type) {
		case 'linear':
			return `y = ${second.toFixed(2)}x + ${first.toFixed(2)}`;
		case 'exponential':
			return `y = ${first.toFixed(2)}e^(${second.toFixed(2)}x)`;
		case 'logarithmic':
			return `y = ${second.toFixed(2)}ln(x) + ${first.toFixed(2)}`;
		case 'power':
			return `y = ${first.toFixed(2)}x^${second.toFixed(2)}`;
		case 'polynomial':
			return fit.coefficients.map((c, i) => `${c.toFixed(2)}x^${i}`).join(' + ');
		default:
			return fit.equation;
	}
}

/**
 * Compute the polyline points (and equation / R-squared) for one trendline
 * over a series' values. Returns an empty point list when the type is
 * unsupported or the data is too sparse to fit.
 */
export function computeTrendlinePoints(
	trendline: PptxChartTrendline,
	values: number[],
	catCount: number,
	layout: PlotLayout,
	range: ValueRange,
	mode: 'line' | 'bar',
): ComputedTrend {
	const fit = fitChartTrendline(
		{
			kind: trendline.trendlineType,
			order: trendline.order,
			period: trendline.period,
			forward: trendline.forward,
			backward: trendline.backward,
			intercept: trendline.intercept,
		},
		values.map((_v, i) => i),
		values,
	);
	if (fit.equation === '') return { points: [], equation: '', rSquared: 0 };
	return {
		points: fit.points.map((point) => ({
			x: xToPixel(point.x, catCount, layout, mode),
			y: valueToYLocal(point.y, range, layout.plotTop, layout.plotBottom),
		})),
		equation: equationText(trendline.trendlineType, fit),
		rSquared: fit.rSquared,
	};
}

/**
 * Map a value to a Y pixel. Kept local (rather than importing `valueToY`) so
 * the trendline maths is independent of any future change to the linear
 * mapping used by the axis renderer  -  they happen to coincide today.
 */
function valueToYLocal(val: number, range: ValueRange, topY: number, bottomY: number): number {
	const usable = bottomY - topY;
	const ratio = (val - range.min) / range.span;
	return range.reverseOrder ? topY + ratio * usable : bottomY - ratio * usable;
}
