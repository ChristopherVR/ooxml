/**
 * Headless chart operations for gradient fills on a whole series
 * (`c:ser/c:spPr/a:gradFill`) or a single data point
 * (`c:dPt/c:spPr/a:gradFill`). Split out of `chart-operations.ts`, which is
 * already well over the repo's 300-line-per-file guideline.
 *
 * Gradient and solid colour are one fill choice, so the last call wins:
 * {@link setChartSeriesGradient} overrides the series colour on save, and
 * `setChartSeriesColor` clears the gradient (likewise
 * {@link setChartDataPointGradient} versus `setChartDataPointFill`).
 *
 * @module sdk/chart-gradient-operations
 */

import type { ChartPptxElement } from '../../types/elements';
import type { ChartGradientInput } from './chart-gradient-input';
import { assertChartGradientSupported, toChartGradientFill } from './chart-gradient-input';
import { ensureDataPoint, removeEmptyDataPoint, validateSeriesIndex } from './chart-operations';

function seriesChartType(element: ChartPptxElement, seriesIndex: number) {
	const data = element.chartData!;
	return data.series[seriesIndex].seriesChartType ?? data.chartType;
}

/**
 * Set (or clear) a gradient fill on a whole series, in a new chart or one
 * loaded from a file. Round-trips to the saved `.pptx` as
 * `c:ser/c:spPr/a:gradFill`, replacing the series' solid fill.
 *
 * Pass `null` to remove the gradient: the series falls back to its solid
 * `color`, or to its automatic theme colour when it has none.
 *
 * @throws {RangeError} If `seriesIndex` is out of bounds or the gradient is
 *   invalid (see {@link toChartGradientFill}).
 * @throws {Error} If the series is drawn as a line-drawn type (line, line3D,
 *   scatter, radar, stock) or a ChartEx type, which have no `c:spPr` area fill.
 *
 * @example
 * ```ts
 * setChartSeriesGradient(chartEl, 0, {
 *   angle: 90, // top to bottom
 *   stops: [
 *     { color: "#60A5FA", position: 0 },
 *     { color: "#1E3A8A", position: 100 },
 *   ],
 * });
 * setChartSeriesGradient(chartEl, 0, null); // back to the solid colour
 * ```
 */
export function setChartSeriesGradient(
	element: ChartPptxElement,
	seriesIndex: number,
	gradient: ChartGradientInput | null,
): void {
	validateSeriesIndex(element, seriesIndex);
	const series = element.chartData!.series[seriesIndex];
	if (gradient === null) {
		delete series.gradientFill;
		return;
	}
	assertChartGradientSupported(seriesChartType(element, seriesIndex), `series ${seriesIndex}`);
	series.gradientFill = toChartGradientFill(gradient);
}

/**
 * Set (or clear) a gradient fill on a single data point (one bar, column or
 * slice), overriding the series fill for that point only. Round-trips to the
 * saved `.pptx` as `c:dPt/c:spPr/a:gradFill` keyed by `c:idx`, replacing any
 * per-point solid fill colour.
 *
 * Pass `null` to remove the per-point gradient (dropping the whole `c:dPt`
 * override when nothing else is set on it).
 *
 * @throws {RangeError} If `seriesIndex` is out of bounds, `pointIndex` is not a
 *   non-negative integer, or the gradient is invalid.
 * @throws {Error} If the series is drawn as a line-drawn or ChartEx type.
 *
 * @example
 * ```ts
 * setChartDataPointGradient(chartEl, 0, 2, {
 *   type: "radial",
 *   stops: [
 *     { color: "#FFFFFF", position: 0 },
 *     { color: "#C00000", position: 100 },
 *   ],
 * });
 * setChartDataPointGradient(chartEl, 0, 2, null); // clear
 * ```
 */
export function setChartDataPointGradient(
	element: ChartPptxElement,
	seriesIndex: number,
	pointIndex: number,
	gradient: ChartGradientInput | null,
): void {
	validateSeriesIndex(element, seriesIndex);
	if (!Number.isInteger(pointIndex) || pointIndex < 0) {
		throw new RangeError(`Data point index must be a non-negative integer, got ${pointIndex}.`);
	}
	const series = element.chartData!.series[seriesIndex];
	if (gradient === null) {
		const dp = series.dataPoints?.find((p) => p.idx === pointIndex);
		if (dp) {
			delete dp.gradientFill;
			removeEmptyDataPoint(series, pointIndex);
		}
		return;
	}
	assertChartGradientSupported(seriesChartType(element, seriesIndex), `series ${seriesIndex}`);
	const gradientFill = toChartGradientFill(gradient);
	const dp = ensureDataPoint(series, pointIndex);
	dp.gradientFill = gradientFill;
	if (dp.spPr?.fillColor !== undefined) {
		const { fillColor: _replaced, ...rest } = dp.spPr;
		dp.spPr = Object.keys(rest).length > 0 ? rest : undefined;
	}
}
