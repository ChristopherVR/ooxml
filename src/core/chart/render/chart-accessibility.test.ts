// A chart is announced as a picture, so its name and a readout of its data are all a reader gets.
import { describe, expect, it } from 'vitest';
import { chartAccessibleName, chartDataDescription } from './chart-accessibility';
import { renderChartSvg } from './chart-svg';
import { chartSummaryView } from './chart-view';
import type { ChartSummary } from './summary';
import { chartThemePalette } from './render-space';

const summary: ChartSummary = {
	chartType: 'column',
	title: 'Revenue by quarter',
	showLegend: true,
	series: [
		{ name: 'Revenue', categories: ['Q1', 'Q2', 'Q3'], values: [12, 18, 9] },
		{ name: 'Cost', categories: ['Q1', 'Q2', 'Q3'], values: [5, 6, null] },
	],
};
const view = (chart: ChartSummary = summary) =>
	chartSummaryView(chart, chartThemePalette(undefined));
const svg = (altText?: string) =>
	renderChartSvg(view(), 480, 300, altText === undefined ? {} : { altText });

describe('a chart’s accessible name', () => {
	it('is the author’s alt text, ahead of the chart title', () => {
		expect(chartAccessibleName(summary, 'Revenue grew every quarter')).toBe(
			'Revenue grew every quarter',
		);
		expect(svg('Revenue grew every quarter')).toContain(
			'<title>Revenue grew every quarter</title>',
		);
	});

	it('falls back to the chart title, and whitespace-only alt text counts as none', () => {
		expect(chartAccessibleName(summary, '   ')).toBe('Revenue by quarter');
		expect(svg()).toContain('<title>Revenue by quarter</title>');
	});

	it('is absent when the chart has neither', () => {
		const { title: _title, ...untitled } = summary;
		expect(chartAccessibleName({})).toBeUndefined();
		expect(renderChartSvg(view(untitled), 480, 300)).not.toContain('<title>');
	});
});

describe('a chart’s data readout', () => {
	it('names each series with its categories and values', () => {
		expect(chartDataDescription(summary)).toBe('Revenue: Q1 12; Q2 18; Q3 9. Cost: Q1 5; Q2 6.');
	});

	it('reaches assistive technology as the painted SVG’s description', () => {
		expect(svg()).toContain('<desc>Revenue: Q1 12; Q2 18; Q3 9. Cost: Q1 5; Q2 6.</desc>');
	});

	it('leaves out a hidden series and a series with no plotted value', () => {
		expect(
			chartDataDescription({
				categories: ['A'],
				series: [
					{ name: 'Shown', values: [1] },
					{ name: 'Hidden', values: [2], hidden: true },
					{ name: 'Empty', values: [null] },
				],
			}),
		).toBe('Shown: A 1.');
	});

	it('is empty when the chart plots nothing, and the SVG then carries no description', () => {
		expect(chartDataDescription({ series: [] })).toBe('');
		expect(renderChartSvg(view({ ...summary, series: [] }), 480, 300)).not.toContain('<desc>');
	});

	it('ends in an ellipsis rather than reading out hundreds of points', () => {
		const readout = chartDataDescription({
			series: [{ name: 'Long', values: Array.from({ length: 50 }, (_, at) => at) }],
		});
		expect(readout.startsWith('Long: 0; 1; 2;')).toBe(true);
		expect(readout).toContain('…');
		expect(readout).not.toContain('; 30;');
	});
});
