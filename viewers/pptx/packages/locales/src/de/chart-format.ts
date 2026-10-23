import { CHART_FORMAT_STRINGS, prefixChartFormatStrings } from 'ooxml-ui/chart-format-locales';

/** Chart Format contextual tab and chart Format task panes. */
export const translations: Record<string, string> = {
	// The shared pane's captions, from its own dictionary in ooxml-ui.
	...prefixChartFormatStrings(CHART_FORMAT_STRINGS.de, 'pptx.'),
	'pptx.ribbon.tab.chartFormat': 'Diagrammformat',
	'pptx.ribbon.groupCurrentSelection': 'Aktuelle Auswahl',
	'pptx.chartFormat.chartElements': 'Diagrammelemente',
	'pptx.chartFormat.formatSelection': 'Auswahl formatieren',
	'pptx.chartFormat.resetToMatchStyle': 'Auf Formatvorlage zurücksetzen',
	'pptx.contextMenu.formatChartArea': 'Diagrammbereich formatieren...',
	'pptx.contextMenu.formatPlotArea': 'Zeichnungsfläche formatieren...',
	'pptx.contextMenu.formatChartTitle': 'Diagrammtitel formatieren...',
	'pptx.contextMenu.formatLegend': 'Legende formatieren...',
	'pptx.contextMenu.formatAxis': 'Achse formatieren...',
	'pptx.contextMenu.formatGridlines': 'Gitternetzlinien formatieren...',
	'pptx.contextMenu.formatDataSeries': 'Datenreihen formatieren...',
	'pptx.contextMenu.formatDataLabels': 'Datenbeschriftungen formatieren...',
};
