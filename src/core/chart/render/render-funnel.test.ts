import { describe, expect, it } from 'vitest';
import { createChartExSpace } from '../ex/create';
import { writeChartExSpace } from '../ex/write';
import { chartExSummary } from './summary-from-chartex';
import { renderChartSummarySvg } from './render-space';
import { computeFunnelGeometry } from '../funnel-geometry';

describe('shared funnel rendering', () => {
	const data = {
		categories: ['A', 'B', 'C', 'D'],
		series: [{ name: 'Sales', values: [100, 0, -50, null] }],
	};
	it('shares centred flat bars, zero width and absolute-value sizing with PowerPoint', () => {
		const bars = computeFunnelGeometry([100, 0, -50], { x: 10, y: 20, width: 200, height: 90 });
		expect(bars.map((b) => [b.x, b.width, b.height])).toEqual([
			[10, 200, 30],
			[110, 0, 30],
			[60, 100, 30],
		]);
	});
	it('reads a saved funnel part and draws values/categories without dropping negative or blank identities', () => {
		const source = createChartExSpace(data, 'funnel');
		const series = source.plotArea.series[0]!;
		series.dataLabels = { visibility: { value: true }, labels: [], hidden: [2] };
		series.dataPoints.push({
			index: 0,
			spPr: { fill: { kind: 'solid', color: { kind: 'srgb', value: 'FF0000', transforms: [] } } },
		});
		const parsed = chartExSummary(writeChartExSpace(source))!;
		expect(parsed.summary?.chartType).toBe('funnel');
		const svg = renderChartSummarySvg(parsed.summary!, { width: 400, height: 300 });
		expect(svg).toContain('data-chart-point="2"');
		expect(svg).not.toContain('data-chart-point="3"');
		expect(svg).toContain('fill="#FF0000"');
		expect(svg).toContain('>100</text>');
		expect(svg).not.toContain('>-50</text>');
		expect(svg).not.toMatch(/NaN|Infinity|not drawn/u);
	});
});
