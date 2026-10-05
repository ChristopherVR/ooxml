/**
 * chart-series-line-style.ts: the stroke a line series is drawn with. Shared by
 * the line chart builder (`chart-cartesian-line-area.ts`) and the line lane of
 * a bar + line combo (`chart-combo-series.ts`), so both pick up the series'
 * own `c:spPr/a:ln` width and dash.
 *
 * @module chart-series-line-style
 */
import type { PptxChartData, PptxChartSeries } from 'pptx-viewer-core';

import { CHART_PX_PER_PT } from './chart-font';
import { resolveChartStyleDefaults } from './chart-style-defaults';
import type { SvgPolyline } from './chart-view-model';
import { buildDashArray } from './connector-dash';

/** PowerPoint's own default line-chart series stroke width when nothing overrides it. */
const DEFAULT_LINE_STROKE_WIDTH = 2.4;

/**
 * Stroke width (px) and dash for one series line.
 *
 * The series' own `a:ln/@w` wins, then the chart style's
 * `cs:dataPointLine`/`cs:dataPoint` width, then 2.4px.
 */
export function seriesLineStroke(
	chartData: PptxChartData,
	series: PptxChartSeries,
): Pick<SvgPolyline, 'strokeWidth' | 'dashArray'> {
	// An authored a:ln width is pt, converted to px. The chart-style width is
	// used as-is.
	const strokeWidth =
		series.lineWidth !== undefined
			? series.lineWidth * CHART_PX_PER_PT
			: (resolveChartStyleDefaults(chartData).seriesLineWidthPt ?? DEFAULT_LINE_STROKE_WIDTH);
	const dashArray = buildDashArray(series.lineDashStyle, strokeWidth);
	return { strokeWidth, ...(dashArray ? { dashArray } : {}) };
}
