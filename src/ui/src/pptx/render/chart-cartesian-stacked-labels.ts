/**
 * chart-cartesian-stacked-labels.ts: the (non-percent) stacked bar/column data
 * labels, split out of `chart-cartesian-bars.ts` to keep that file within the
 * repo's ~300-LOC limit.
 *
 * @module chart-cartesian-stacked-labels
 */
import type { PptxChartData, PptxChartSeries } from 'ooxml-core/pptx';

import { resolveBarLabelPlacement } from './chart-data-label-anchor';
import {
	buildDataLabelText,
	dataLabelFontOverride,
	resolveDataLabelTextStyle,
} from './chart-data-label-text';
import { DEFAULT_CHART_DATA_LABEL_PX } from './chart-font';
import type { BarRect, PlotLayout, SvgText } from './chart-view-model';

/**
 * Push the stacked column data labels, one per drawn segment. Each label is
 * placed on its own segment's rect, which `computeStackedBarRects` has
 * already clipped to the value axis: a segment that an authored `c:min` /
 * `c:max` cuts away entirely has no rect and so no label, and a partly cut
 * segment is labelled on its visible part (its centre for `ctr`, the
 * default when no `c:dLblPos` is authored). `c:dLblPos` (ctr/inBase/inEnd/outEnd) and a per-point
 * `c:dLbl/c:layout` drag go through the same `resolveBarLabelPlacement`
 * pipeline the clustered path uses. A zero value draws no segment and no
 * label, as on the horizontal stacked bar.
 */
export function pushStackedBarLabels(
	chartData: PptxChartData,
	series: ReadonlyArray<PptxChartSeries>,
	sourceIndices: ReadonlyArray<number>,
	rects: ReadonlyArray<BarRect>,
	layout: PlotLayout,
	dataLabels: SvgText[],
): void {
	for (const rect of rects) {
		const entry = rect.seriesIndex === undefined ? undefined : series[rect.seriesIndex];
		if (!entry || rect.pointIndex === undefined) {
			continue;
		}
		const sourceIndex = sourceIndices[rect.pointIndex] ?? rect.pointIndex,
			val = entry.values[sourceIndex] ?? 0,
			label = buildDataLabelText({ chartData, series: entry, pointIndex: sourceIndex, value: val });
		if (label === undefined) {
			continue;
		}
		const anchor = resolveBarLabelPlacement(
			chartData,
			entry,
			sourceIndex,
			{ x: rect.x, y: rect.y, width: rect.w, height: rect.h },
			val,
			'vertical',
			{ width: layout.svgWidth, height: layout.svgHeight },
			'ctr',
		);
		dataLabels.push({
			kind: 'text',
			x: anchor.x,
			y: anchor.y,
			text: label.text,
			fontSize: DEFAULT_CHART_DATA_LABEL_PX,
			fill: label.color ?? '#334155',
			textAnchor: anchor.textAnchor,
			...(anchor.dominantBaseline ? { dominantBaseline: anchor.dominantBaseline } : {}),
			...dataLabelFontOverride(resolveDataLabelTextStyle(chartData, entry, sourceIndex)),
		});
	}
}
