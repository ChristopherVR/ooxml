import { CHART_FORMAT_STRINGS, prefixChartFormatStrings } from 'ooxml-ui/chart-format-locales';

/** Chart Format contextual tab and chart Format task panes. */
export const translations: Record<string, string> = {
	// The shared pane's captions, from its own dictionary in ooxml-ui.
	...prefixChartFormatStrings(CHART_FORMAT_STRINGS.ja, 'pptx.'),
	'pptx.ribbon.tab.chartFormat': '書式',
	'pptx.ribbon.groupCurrentSelection': '現在の選択範囲',
	'pptx.chartFormat.chartElements': 'グラフ要素',
	'pptx.chartFormat.formatSelection': '選択対象の書式設定',
	'pptx.chartFormat.resetToMatchStyle': 'リセットしてスタイルに合わせる',
	'pptx.contextMenu.formatChartArea': 'グラフ エリアの書式設定...',
	'pptx.contextMenu.formatPlotArea': 'プロット エリアの書式設定...',
	'pptx.contextMenu.formatChartTitle': 'グラフ タイトルの書式設定...',
	'pptx.contextMenu.formatLegend': '凡例の書式設定...',
	'pptx.contextMenu.formatAxis': '軸の書式設定...',
	'pptx.contextMenu.formatGridlines': '目盛線の書式設定...',
	'pptx.contextMenu.formatDataSeries': 'データ系列の書式設定...',
	'pptx.contextMenu.formatDataLabels': 'データ ラベルの書式設定...',
};
