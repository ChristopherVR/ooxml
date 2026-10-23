import { CHART_FORMAT_STRINGS, prefixChartFormatStrings } from 'ooxml-ui/chart-format-locales';

/** Chart Format contextual tab and chart Format task panes. */
export const translations: Record<string, string> = {
	// The shared pane's captions, from its own dictionary in ooxml-ui.
	...prefixChartFormatStrings(CHART_FORMAT_STRINGS.fr, 'pptx.'),
	'pptx.ribbon.tab.chartFormat': 'Format du graphique',
	'pptx.ribbon.groupCurrentSelection': 'Sélection active',
	'pptx.chartFormat.chartElements': 'Éléments de graphique',
	'pptx.chartFormat.formatSelection': 'Mettre en forme la sélection',
	'pptx.chartFormat.resetToMatchStyle': 'Réinitialiser pour correspondre au style',
	'pptx.contextMenu.formatChartArea': 'Mettre en forme la zone de graphique...',
	'pptx.contextMenu.formatPlotArea': 'Mettre en forme la zone de traçage...',
	'pptx.contextMenu.formatChartTitle': 'Mettre en forme le titre du graphique...',
	'pptx.contextMenu.formatLegend': 'Mettre en forme la légende...',
	'pptx.contextMenu.formatAxis': 'Mettre en forme l’axe...',
	'pptx.contextMenu.formatGridlines': 'Mettre en forme le quadrillage...',
	'pptx.contextMenu.formatDataSeries': 'Mettre en forme des séries de données...',
	'pptx.contextMenu.formatDataLabels': 'Mettre en forme les étiquettes de données...',
};
