/**
 * The point paint a series colour edit clears, as the Format pane's colour picker
 * does and as `chart/edit` `setChartSeriesColor` writes it: each point's own fill
 * (`c:dPt/c:spPr` fill), or its own line for a series drawn as a line (line,
 * scatter, stock, unfilled radar). A point override with something else of its
 * own (a marker, an explosion) stays; one with nothing left is removed.
 *
 * @module sdk/chart-series-point-paint
 */

import type { PptxChartData, PptxChartSeries } from '../../types/chart';
import { isLineDrawnChartType } from '../../utils/chart-container-type-map';

/** Whether the series is coloured by its line rather than its fill. */
function paintsLine(data: PptxChartData, series: PptxChartSeries): boolean {
	const type = series.seriesChartType ?? data.chartType;
	if (type === 'radar') return data.radarStyle !== 'filled';
	return isLineDrawnChartType(type);
}

/**
 * Clear the fill (or, for a line-drawn series, the line) of every point of `series`. A point left
 * with no override of its own is removed, so it takes the series colour.
 */
export function clearChartSeriesPointPaint(data: PptxChartData, series: PptxChartSeries): void {
	if (!series.dataPoints) return;
	const line = paintsLine(data, series);
	for (const point of series.dataPoints) {
		if (!line) delete point.gradientFill;
		if (!point.spPr) continue;
		const keys = line
			? (['strokeColor', 'strokeWidth', 'strokeDashStyle', 'lineNoFill'] as const)
			: (['fillColor'] as const);
		for (const key of keys) delete point.spPr[key];
		if (Object.keys(point.spPr).length === 0) delete point.spPr;
	}
	series.dataPoints = series.dataPoints.filter((point) =>
		Object.entries(point).some(([name, value]) => name !== 'idx' && value !== undefined),
	);
}
