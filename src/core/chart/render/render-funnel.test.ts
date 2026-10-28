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
		expect(bars.map((b) => [b.x, b.width])).toEqual([
			[10, 200],
			[110, 0],
			[60, 100],
		]);
		expect(bars[0]!.height).toBeCloseTo(30 / 1.33);
		expect(bars[0]!.y + bars[0]!.height / 2).toBeCloseTo(35);
		expect(bars[1]!.y - bars[0]!.y).toBeCloseTo(30);
	});
	it('honors explicit touching bars and a gap equal to the bar height', () => {
		const plot = { x: 0, y: 0, width: 200, height: 100 };
		const touching = computeFunnelGeometry([100, 50], plot, 0);
		expect(touching[0]!.y + touching[0]!.height).toBe(touching[1]!.y);
		const spaced = computeFunnelGeometry([100, 50], plot, 100);
		expect(spaced.map((bar) => [bar.y, bar.height])).toEqual([
			[12.5, 25],
			[62.5, 25],
		]);
	});
	it('carries authored gap width through the shared painter', () => {
		const summary = chartExSummary(writeChartExSpace(createChartExSpace(data, 'funnel')))!.summary!;
		const paint = (gap: number) =>
			renderChartSummarySvg({ ...summary, funnelGapWidth: gap }, { width: 400, height: 300 });
		const heights = (svg: string) =>
			[
				...svg.matchAll(
					/<g data-chart-series="0" data-chart-point="\d+"><rect [^>]*height="([\d.]+)"/g,
				),
			].map((match) => Number(match[1]));
		expect(heights(paint(0)).length).toBeGreaterThan(0);
		expect(heights(paint(0))[0]).toBeCloseTo(heights(paint(100))[0]! * 2, 1);
	});
	it('does not invent value labels when the source declares none', () => {
		const summary = chartExSummary(writeChartExSpace(createChartExSpace(data, 'funnel')))!.summary!;
		const svg = renderChartSummarySvg(summary, { width: 400, height: 300 });
		expect(svg).not.toContain('>100</text>');
		expect(svg).toContain('>A</text>');
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
