/**
 * The chart entry of the canvas context menu: PowerPoint's "Format Chart
 * Area...", "Format Data Series...", "Format Axis..." and so on, named after
 * the chart element the menu was opened on.
 *
 * @module render/chart-format/chart-format-context-menu
 */
import type { PptxChartFormatTargetKey } from 'ooxml-core/pptx';
import { parseChartFormatTargetKey } from 'ooxml-core/pptx';

const CHART_AREA_KEY = 'pptx.contextMenu.formatChartArea';

const KEYS: Readonly<Record<string, string>> = {
	chartArea: CHART_AREA_KEY,
	plotArea: 'pptx.contextMenu.formatPlotArea',
	title: 'pptx.contextMenu.formatChartTitle',
	legend: 'pptx.contextMenu.formatLegend',
	axis: 'pptx.contextMenu.formatAxis',
	majorGridlines: 'pptx.contextMenu.formatGridlines',
	minorGridlines: 'pptx.contextMenu.formatGridlines',
	series: 'pptx.contextMenu.formatDataSeries',
	dataLabels: 'pptx.contextMenu.formatDataLabels',
};

/** The i18n key of the context-menu entry for the chart element `targetKey`. */
export function chartFormatContextMenuLabelKey(
	targetKey: PptxChartFormatTargetKey | null | undefined,
): string {
	const kind = targetKey ? parseChartFormatTargetKey(targetKey)?.kind : undefined;
	return (kind && KEYS[kind]) || CHART_AREA_KEY;
}
