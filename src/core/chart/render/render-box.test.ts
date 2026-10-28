import { describe, expect, it } from 'vitest';
import { createChartExSpace } from '../ex/create';
import { writeChartExSpace } from '../ex/write';
import { chartExSummary } from './summary-from-chartex';
import { renderChartSummarySvg } from './render-space';
import { boxWhiskerDistribution } from '../box-distribution';

describe('shared box and whisker rendering', () => {
	it('shares inclusive/exclusive quartiles and admitted outlier fences with PPTX', () => {
		const values = [1, 2, 3, 4, 100];
		expect(boxWhiskerDistribution(values, 'inclusive', true)).toMatchObject({
			whiskerMin: 1,
			whiskerMax: 4,
			stats: { q1: 2, median: 3, q3: 4 },
			mean: 22,
		});
		expect(boxWhiskerDistribution(values, 'exclusive', false)?.whiskerMax).toBe(100);
		expect(boxWhiskerDistribution([1])).toBeUndefined();
	});
	it('keeps every series and groups repeated category observations into separate boxes', () => {
		const space = createChartExSpace(
			{
				categories: ['A', 'A', 'A', 'B', 'B', 'B'],
				series: [
					{ name: 'One', values: [1, 2, 3, 4, 5, 6] },
					{ name: 'Two', values: [2, 3, 4, 5, 6, 7] },
				],
			},
			'boxWhisker',
		);
		const summary = chartExSummary(writeChartExSpace(space))!.summary!;
		expect(summary.series).toHaveLength(2);
		const svg = renderChartSummarySvg(summary, { width: 500, height: 300 });
		expect(svg.match(/data-chart-part="box"/gu)).toHaveLength(4);
		expect(svg).toContain('data-chart-series="1"');
		expect(svg).not.toMatch(/NaN|Infinity|not drawn/u);
	});
	it('retains observation identities, mean/outlier visibility and value-axis bounds', () => {
		const space = createChartExSpace(
			{
				categories: ['A', 'A', 'A', 'A', 'A'],
				series: [{ name: 'One', values: [1, 2, 3, 4, 100] }],
			},
			'boxWhisker',
		);
		space.plotArea.series[0]!.layoutProperties = {
			statistics: { quartileMethod: 'inclusive' },
			visibility: { meanMarker: true, meanLine: true, outliers: true },
		};
		const axis = space.plotArea.axes.find((a) => a.scaling?.kind === 'val')!;
		axis.scaling = { kind: 'val', min: -10, max: 120 };
		const before = writeChartExSpace(space);
		const summary = chartExSummary(before)!.summary!;
		const svg = renderChartSummarySvg(summary, { width: 500, height: 300 });
		expect(summary.axes?.find((a) => a.kind === 'val')?.scaling).toEqual({ min: -10, max: 120 });
		expect(svg).toContain('data-chart-point="4" data-chart-part="outlier"');
		expect(svg).toContain('data-chart-part="mean"');
		expect(writeChartExSpace(space)).toBe(before);
	});
	it('does not draw hidden series or invent a box from a single observation', () => {
		const space = createChartExSpace(
			{
				categories: ['A', 'A', 'B'],
				series: [
					{ name: 'One', values: [1, 2, 9] },
					{ name: 'Two', values: [3, 4, 10] },
				],
			},
			'boxWhisker',
		);
		space.plotArea.series[0]!.hidden = true;
		const svg = renderChartSummarySvg(chartExSummary(writeChartExSpace(space))!.summary!, {
			width: 400,
			height: 300,
		});
		expect(svg.match(/data-chart-part="box"/gu)).toHaveLength(1);
		expect(svg).toContain('data-chart-series="1"');
	});
});
