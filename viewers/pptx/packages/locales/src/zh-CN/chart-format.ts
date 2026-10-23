import { CHART_FORMAT_STRINGS, prefixChartFormatStrings } from 'ooxml-ui/chart-format-locales';

/** Chart Format contextual tab and chart Format task panes. */
export const translations: Record<string, string> = {
	// The shared pane's captions, from its own dictionary in ooxml-ui.
	...prefixChartFormatStrings(CHART_FORMAT_STRINGS['zh-CN'], 'pptx.'),
	'pptx.ribbon.tab.chartFormat': '图表格式',
	'pptx.ribbon.groupCurrentSelection': '当前所选内容',
	'pptx.chartFormat.chartElements': '图表元素',
	'pptx.chartFormat.formatSelection': '设置所选内容格式',
	'pptx.chartFormat.resetToMatchStyle': '重设以匹配样式',
	'pptx.contextMenu.formatChartArea': '设置图表区格式...',
	'pptx.contextMenu.formatPlotArea': '设置绘图区格式...',
	'pptx.contextMenu.formatChartTitle': '设置图表标题格式...',
	'pptx.contextMenu.formatLegend': '设置图例格式...',
	'pptx.contextMenu.formatAxis': '设置坐标轴格式...',
	'pptx.contextMenu.formatGridlines': '设置网格线格式...',
	'pptx.contextMenu.formatDataSeries': '设置数据系列格式...',
	'pptx.contextMenu.formatDataLabels': '设置数据标签格式...',
};
