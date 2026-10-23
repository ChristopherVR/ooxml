// Excel's source for the shared chart Format pane (`office-ui-chart-format-pane`): what an xlsx
// `ChartObject` reads as. The chart summary is not a `ChartSpace`, so the pane gets the elements
// and options Excel's `ChartPatch` path writes: the chart area, plot area, title and legend fills,
// series fills (bar and column charts), series spacing and the legend position. Fills are drawn by
// Excel's own fill controls in the pane's `fill` slot (preset gradients, stops, directions).
import { chartBarSpacing, type ChartObject, type ChartViewModel } from 'ooxml-core/xlsx';
import type { ChartFormatTarget } from 'ooxml-core/chart';
import {
	chartFormatTargetOptions,
	type ChartFormatElementView,
	type ChartFormatSource,
} from '../chart/format-pane-api';

/** The slot Excel's fill controls go into. */
export const XLSX_CHART_FILL_SLOT = 'fill';

/** Bar and column charts: the series Excel formats (spacing and fill). */
export const formatsSeries = (chart: ChartObject): boolean =>
	chart.chartType === 'bar' || chart.chartType === 'column';

const ELEMENT: ChartFormatElementView = {
	support: {
		fill: true,
		gradient: true,
		pattern: false,
		picture: false,
		threeD: false,
		text: false,
		line: false,
		effects: false,
	},
	fill: { kind: 'automatic' },
	line: { kind: 'automatic' },
};

/** The Format pane source of `chart`, drawn as `view`. */
export function xlsxChartFormatSource(chart: ChartObject, view: ChartViewModel): ChartFormatSource {
	const targets: ChartFormatTarget[] = [{ kind: 'chartArea' }, { kind: 'plotArea' }];
	if (chart.title) targets.push({ kind: 'title' });
	if (chart.showLegend) targets.push({ kind: 'legend' });
	if (formatsSeries(chart))
		chart.series.forEach((_, seriesIndex) => targets.push({ kind: 'series', seriesIndex }));
	const spacing = chartBarSpacing(chart);
	return {
		targets: chartFormatTargetOptions(targets, {
			axis: () => undefined,
			seriesName: (index) => view.series[index]?.name || `Series ${index + 1}`,
		}),
		element: () => ELEMENT,
		seriesOptions: () =>
			formatsSeries(chart)
				? {
						chartType: 'bar',
						overlap: spacing.barOverlap,
						gapWidth: spacing.barGapWidth,
						sliders: true,
					}
				: undefined,
		legendOptions: () =>
			chart.showLegend
				? { position: chart.legendPosition ?? 'r', overlaySupported: false }
				: undefined,
		slots: { fill: XLSX_CHART_FILL_SLOT },
		flat: true,
		textOptions: false,
		picker: { key: 'chartFormat.chartElement', fallback: 'Chart element' },
	};
}
