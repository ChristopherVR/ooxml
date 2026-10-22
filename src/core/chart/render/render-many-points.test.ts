import { describe, expect, it } from 'vitest';
import { chartSummaryView } from './chart-view';
import { chartThemePalette, renderChartSummarySvg } from './render-space';
import type { ChartSummary, ChartSummaryType } from './summary';

// More points than a call can pass as arguments, so a spread of them would throw a RangeError.
const COUNT = 200_000;
const values = Array.from({ length: COUNT }, (_, i) => (i % 97) + 1);
const labels = values.map((_, i) => `P${i + 1}`);
const theme = chartThemePalette(undefined);
const size = { width: 480, height: 300 };
const summary = (
	chartType: ChartSummaryType,
	categories: (string | number)[],
	showLegend = false,
): ChartSummary => ({ chartType, showLegend, series: [{ categories, values }] });

describe('charts with very many points', () => {
	it('scales the x axis of a scatter chart over every x value', () => {
		const x = values.map((_, i) => ((i * 7919) % COUNT) + 1);
		const view = chartSummaryView(summary('scatter', x), theme);
		expect(view.xAxis?.min).toBeLessThanOrEqual(1);
		expect(view.xAxis?.max).toBeGreaterThanOrEqual(COUNT);
	});

	it('scales the value axis of a waterfall over every step', () => {
		const total = values.reduce((sum, value) => sum + value, 0);
		const view = chartSummaryView(summary('waterfall', labels), theme);
		expect(view.valueAxis?.max).toBeGreaterThanOrEqual(total);
	});

	it('measures the category labels of a column chart and draws every bar', () => {
		const svg = renderChartSummarySvg(summary('column', labels), size);
		expect(svg).toContain(`data-chart-series="0" data-chart-point="${COUNT - 1}"`);
	});

	it('sizes the legend of a pie chart that lists every point', () => {
		const svg = renderChartSummarySvg(summary('pie', labels, true), size);
		expect(svg).toContain('>P1<');
	});
});
