import { describe, expect, it } from 'vitest';
import type { ChartPlotGroup, ChartSpaceSeries } from '../model';
import { chartMarkerRadius, chartSeriesMarkerSvg } from './chart-svg-markers';
import { seriesMarkerViews, summaryMarkers } from './chart-view-markers';
import { scatter } from './chart-svg-xy';
import { radarSvg } from './chart-svg-radial';
import type { ChartSeriesView, ChartViewModel } from './chart-view';

const group = (
	kind: ChartPlotGroup['kind'],
	overrides: Partial<ChartPlotGroup> = {},
): ChartPlotGroup => ({
	kind,
	element: `${kind}Chart`,
	is3D: false,
	series: [],
	axisIds: [],
	...overrides,
});
const source = (overrides: Partial<ChartSpaceSeries> = {}): ChartSpaceSeries => ({
	dataPoints: [],
	...overrides,
});
const color = () => '#123456';

describe('neutral OOXML markers', () => {
	it('uses the series OOXML index for automatic symbols', () => {
		expect(summaryMarkers(source({ index: 2, marker: {} }), group('line'), 0).marker?.symbol).toBe(
			'triangle',
		);
		expect(summaryMarkers(source(), group('line'), 0).marker?.symbol).toBe('none');
		expect(summaryMarkers(source(), group('line', { marker: true }), 0).marker?.symbol).toBe(
			'diamond',
		);
	});

	it.each(['line', 'smooth', 'none'])('hides scatter markers for %s style', (scatterStyle) => {
		expect(summaryMarkers(source(), group('scatter', { scatterStyle }), 0).marker?.symbol).toBe(
			'none',
		);
	});

	it('cascades sparse point symbol, size, fill and outline independently', () => {
		const summary = summaryMarkers(
			source({
				marker: {
					symbol: 'circle',
					size: 12,
					spPr: {
						fill: { kind: 'none' },
						line: {
							widthEmu: 19050,
							fill: { kind: 'solid', color: { kind: 'srgb', value: '123456', transforms: [] } },
						},
					},
				},
				dataPoints: [
					{ index: 7, marker: { symbol: 'star' } },
					{ index: 2, marker: { symbol: 'none' } },
				],
			}),
			group('line'),
			0,
		);
		const view = { color: '#999', ...seriesMarkerViews(summary, color) };
		expect(view.pointMarkers?.[7]).toMatchObject({
			symbol: 'star',
			size: 12,
			fill: 'none',
			stroke: '#123456',
			strokeWidth: 2,
		});
		expect(chartSeriesMarkerSvg(view, 7, 10, 20)).toContain('<polygon');
		expect(chartSeriesMarkerSvg(view, 7, 10, 20)).toContain('stroke-width="2"');
		expect(chartMarkerRadius(view, 7)).toBe(6);
		expect(chartSeriesMarkerSvg(view, 2, 10, 20)).toBe('');
		expect(chartMarkerRadius(view, 2)).toBe(0);
	});

	it('allows a point-only automatic marker in a series without markers', () => {
		const summary = summaryMarkers(
			source({ index: 1, dataPoints: [{ index: 3, marker: {} }] }),
			group('line'),
			0,
		);
		expect(summary.marker?.symbol).toBe('none');
		expect(summary.pointMarkers?.[3]?.symbol).toBe('square');
	});

	it('escapes marker paint and filter attributes', () => {
		const svg = chartSeriesMarkerSvg({ color: '"<', shadowFilter: '"<' }, 0, 1, 2);
		expect(svg).toContain('fill="&quot;&lt;"');
		expect(svg).toContain('filter="&quot;&lt;"');
	});

	it('paints the shared marker geometry in scatter and radar', () => {
		const series: ChartSeriesView = Object.assign(
			{ name: 'S', values: [2], xValues: [1], color: '#123456' },
			{ marker: { symbol: 'square', size: 12 } },
		);
		const model: ChartViewModel = {
			type: 'scatter',
			grouping: 'standard',
			categories: ['A'],
			series: [series],
			showLegend: false,
			legendPosition: 'r',
			horizontal: false,
			supported: true,
			xAxis: { min: 0, max: 3, ticks: [0, 1, 2, 3], majorUnit: 1 },
			valueAxis: { min: 0, max: 3, ticks: [0, 1, 2, 3], majorUnit: 1, percent: false },
		};
		const out: string[] = [];
		scatter(model, { x: 0, y: 0, w: 300, h: 200 }, out);
		expect(out.join('')).toContain('width="12" height="12" fill="#123456"');
		expect(radarSvg({ ...model, type: 'radar' }, { x: 0, y: 0, w: 300, h: 200 })).toContain(
			'width="12" height="12" fill="#123456"',
		);
	});
});
