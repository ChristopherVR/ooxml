import { CHART_FORMAT_STRINGS, prefixChartFormatStrings } from 'ooxml-ui/chart-format-locales';

/** Chart Format contextual tab and chart Format task panes. */
export const translations: Record<string, string> = {
	// The shared pane's captions, from its own dictionary in ooxml-ui.
	...prefixChartFormatStrings(CHART_FORMAT_STRINGS.ko, 'pptx.'),
	'pptx.ribbon.tab.chartFormat': '서식',
	'pptx.ribbon.groupCurrentSelection': '현재 선택 영역',
	'pptx.chartFormat.chartElements': '차트 요소',
	'pptx.chartFormat.formatSelection': '선택 영역 서식',
	'pptx.chartFormat.resetToMatchStyle': '스타일에 맞게 다시 설정',
	'pptx.contextMenu.formatChartArea': '차트 영역 서식...',
	'pptx.contextMenu.formatPlotArea': '그림 영역 서식...',
	'pptx.contextMenu.formatChartTitle': '차트 제목 서식...',
	'pptx.contextMenu.formatLegend': '범례 서식...',
	'pptx.contextMenu.formatAxis': '축 서식...',
	'pptx.contextMenu.formatGridlines': '눈금선 서식...',
	'pptx.contextMenu.formatDataSeries': '데이터 계열 서식...',
	'pptx.contextMenu.formatDataLabels': '데이터 레이블 서식...',
};
