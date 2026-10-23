import { CHART_FORMAT_STRINGS, prefixChartFormatStrings } from 'ooxml-ui/chart-format-locales';

/** Chart Format contextual tab and chart Format task panes. */
export const translations: Record<string, string> = {
	// The shared pane's captions, from its own dictionary in ooxml-ui.
	...prefixChartFormatStrings(CHART_FORMAT_STRINGS.es, 'pptx.'),
	'pptx.ribbon.tab.chartFormat': 'Formato de gráfico',
	'pptx.ribbon.groupCurrentSelection': 'Selección actual',
	'pptx.chartFormat.chartElements': 'Elementos de gráfico',
	'pptx.chartFormat.formatSelection': 'Dar formato a la selección',
	'pptx.chartFormat.resetToMatchStyle': 'Restablecer para coincidir con el estilo',
	'pptx.contextMenu.formatChartArea': 'Dar formato al área del gráfico...',
	'pptx.contextMenu.formatPlotArea': 'Dar formato al área de trazado...',
	'pptx.contextMenu.formatChartTitle': 'Dar formato al título del gráfico...',
	'pptx.contextMenu.formatLegend': 'Dar formato a la leyenda...',
	'pptx.contextMenu.formatAxis': 'Dar formato al eje...',
	'pptx.contextMenu.formatGridlines': 'Dar formato a líneas de cuadrícula...',
	'pptx.contextMenu.formatDataSeries': 'Dar formato a serie de datos...',
	'pptx.contextMenu.formatDataLabels': 'Dar formato a etiquetas de datos...',
};
