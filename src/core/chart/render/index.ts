/** The DOM-free chart painter: a chart summary or a parsed chart part to a standalone SVG. */
export * from './summary';
export * from './summary-from-space';
export * from './render-space';
export {
	chartSummaryView,
	categoryTotals,
	type ChartRefValue,
	type ChartSeriesView,
	type ChartViewModel,
	type ChartViewOptions,
	type EvaluateRef,
	type ValueAxisView,
} from './chart-view';
export { renderChartSvg, type ChartSvgOptions } from './chart-svg';
export { autoSeriesColor, chartColorScheme, modulateLuminance } from './chart-colors';
export { THEME_SLOTS, DEFAULT_THEME_COLORS, themeColor } from './theme-palette';
