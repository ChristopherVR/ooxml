/**
 * The text a chart's title shows, including PowerPoint's AUTO title.
 *
 * The rule (single series name, "Chart Title" otherwise, chartex families never take the series
 * name) lives in `ooxml-core/chart` (`auto-title`); this module applies it to `PptxChartData`.
 * Every chart family's view model resolves its title through here, so the 2D chart and the 3D
 * chart's chrome overlay agree.
 *
 * @module chart-auto-title
 */
import { CHART_AUTO_TITLE, impliesAutoTitle, resolveAutoTitleText } from 'ooxml-core/chart';
import type { PptxChartData } from 'ooxml-core/pptx';

export { CHART_AUTO_TITLE };

const autoTitleInput = (chartData: PptxChartData) => ({
	chartType: chartData.chartType,
	seriesNames: chartData.series.map((series) => series.name ?? ''),
	title: chartData.title,
	hasTitle: chartData.style?.hasTitle,
	autoTitleDeleted: chartData.chartChrome?.autoTitleDeleted,
});

/**
 * A classic chart with NO `c:title` element still gets an auto title when it plots exactly one
 * series and `c:autoTitleDeleted` is not set (COM-verified, charts-com.pptx slides 5, 6 and 24).
 * Returns the chart with `style.hasTitle` switched on in that case, so the layout reserves the
 * title band; otherwise the chart itself.
 */
export function withAutoTitle(chartData: PptxChartData): PptxChartData {
	return impliesAutoTitle(autoTitleInput(chartData))
		? { ...chartData, style: { ...chartData.style, hasTitle: true } }
		: chartData;
}

/** The title text to draw, or `undefined` when the chart shows no title. */
export function resolveChartTitleText(chartData: PptxChartData): string | undefined {
	return resolveAutoTitleText(autoTitleInput(chartData));
}
