/**
 * Sample charts for the chart Format tests: a column chart with every element
 * a pane can target, a line chart, a pie and a doughnut.
 */
import type { PptxChartData } from 'ooxml-core/pptx';

export function columnChart(): PptxChartData {
	return {
		chartType: 'bar',
		title: 'Sales',
		categories: ['Q1', 'Q2', 'Q3'],
		series: [
			{ name: 'East', values: [10, 20, 30], color: '#4472C4' },
			{ name: 'West', values: [15, 25, 5], color: '#ED7D31' },
		],
		grouping: 'clustered',
		style: { hasTitle: true, hasLegend: true, legendPosition: 'r', hasDataLabels: true },
		axes: [
			{ axisType: 'catAx', axisId: 1, axPos: 'b' },
			{ axisType: 'valAx', axisId: 2, axPos: 'l', majorGridlines: true, minorGridlines: true },
			{ axisType: 'valAx', axisId: 3, axPos: 'r', logScale: true, logBase: 10 },
		],
	};
}

export const lineChart = (): PptxChartData => ({ ...columnChart(), chartType: 'line' });
export const pieChart = (): PptxChartData => ({ ...columnChart(), chartType: 'pie', axes: [] });
export const doughnutChart = (): PptxChartData => ({ ...pieChart(), chartType: 'doughnut' });
