import { chartFormatEn } from '../../chart/format-pane-locales/en';
import { prefixChartFormatStrings } from '../../chart/format-pane-locales/index';

/** English copy for the chart Format task panes and the Chart Format contextual tab. */
export const chartFormatTranslationsEn: Record<string, string> = {
	// The shared pane's captions, from its own dictionary (`chart/format-pane-locales`).
	...prefixChartFormatStrings(chartFormatEn, 'pptx.'),
	'pptx.ribbon.tab.chartFormat': 'Chart Format',
	'pptx.ribbon.groupCurrentSelection': 'Current Selection',
	'pptx.chartFormat.chartElements': 'Chart Elements',
	'pptx.chartFormat.formatSelection': 'Format Selection',
	'pptx.chartFormat.resetToMatchStyle': 'Reset to Match Style',
	'pptx.contextMenu.formatChartArea': 'Format Chart Area...',
	'pptx.contextMenu.formatPlotArea': 'Format Plot Area...',
	'pptx.contextMenu.formatChartTitle': 'Format Chart Title...',
	'pptx.contextMenu.formatLegend': 'Format Legend...',
	'pptx.contextMenu.formatAxis': 'Format Axis...',
	'pptx.contextMenu.formatGridlines': 'Format Gridlines...',
	'pptx.contextMenu.formatDataSeries': 'Format Data Series...',
	'pptx.contextMenu.formatDataLabels': 'Format Data Labels...',
};
