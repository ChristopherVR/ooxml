import type { ChartPptxElement, PptxChartData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	clipPolygonToBand,
	clipPolylineToBand,
	clippedSeriesLine,
	visibleSpanCentre,
} from './chart-plot-clip';
import { buildChartViewModel } from './chart-view-model';
import type { SvgPrimitive } from './chart-view-model';

const band = { top: 10, bottom: 110 };

describe('clipPolylineToBand', () => {
	it('returns a line inside the band unchanged', () => {
		const points = [
			{ x: 0, y: 20 },
			{ x: 10, y: 100 },
		];
		const runs = clipPolylineToBand(points, band);
		expect(runs).toHaveLength(1);
		expect(runs[0]?.[0]).toBe(points[0]);
		expect(runs[0]?.[1]).toBe(points[1]);
	});

	it('cuts a line at the top edge and starts a new run where it comes back', () => {
		const runs = clipPolylineToBand(
			[
				{ x: 0, y: 30 },
				{ x: 10, y: -10 },
				{ x: 20, y: 30 },
				{ x: 30, y: 50 },
			],
			band,
		);
		expect(runs).toStrictEqual([
			[
				{ x: 0, y: 30 },
				{ x: 5, y: 10 },
			],
			[
				{ x: 15, y: 10 },
				{ x: 20, y: 30 },
				{ x: 30, y: 50 },
			],
		]);
	});

	it('keeps the part of a segment that crosses the whole band', () => {
		expect(
			clipPolylineToBand(
				[
					{ x: 0, y: 0 },
					{ x: 120, y: 120 },
				],
				band,
			),
		).toStrictEqual([
			[
				{ x: 10, y: 10 },
				{ x: 110, y: 110 },
			],
		]);
	});

	it('drops a line wholly outside the band', () => {
		expect(
			clipPolylineToBand(
				[
					{ x: 0, y: 0 },
					{ x: 10, y: 5 },
				],
				band,
			),
		).toStrictEqual([]);
	});
});

describe('clipPolygonToBand', () => {
	it('returns a polygon inside the band unchanged', () => {
		const polygon = [
			{ x: 0, y: 20 },
			{ x: 10, y: 20 },
			{ x: 10, y: 110 },
			{ x: 0, y: 110 },
		];
		expect(clipPolygonToBand(polygon, band)).toStrictEqual(polygon);
	});

	it('cuts a polygon at both edges', () => {
		const clipped = clipPolygonToBand(
			[
				{ x: 0, y: 0 },
				{ x: 10, y: 0 },
				{ x: 10, y: 200 },
				{ x: 0, y: 200 },
			],
			band,
		);
		expect(Math.min(...clipped.map((p) => p.y))).toBe(10);
		expect(Math.max(...clipped.map((p) => p.y))).toBe(110);
		expect(clipped).toHaveLength(4);
	});

	it('drops a polygon wholly outside the band', () => {
		expect(
			clipPolygonToBand(
				[
					{ x: 0, y: 0 },
					{ x: 10, y: 0 },
					{ x: 10, y: 5 },
				],
				band,
			),
		).toStrictEqual([]);
	});
});

describe('visibleSpanCentre', () => {
	it('centres on the visible part of a span', () => {
		expect(visibleSpanCentre(50, 90, band)).toBe(70);
		expect(visibleSpanCentre(-30, 50, band)).toBe(30);
		expect(visibleSpanCentre(-30, 0, band)).toBeUndefined();
	});
});

describe('clippedSeriesLine', () => {
	const style = { stroke: '#000', strokeWidth: 1, fill: 'none' };
	const crossing = [
		{ x: 0, y: 50 },
		{ x: 10, y: -50 },
		{ x: 20, y: 50 },
	];

	it('splits a clipped straight line into one polyline per run', () => {
		const lines = clippedSeriesLine(crossing, false, style, band);
		expect(lines.map((line) => line.kind)).toStrictEqual(['polyline', 'polyline']);
	});

	it('keeps a clipped smooth line as one path inside the band', () => {
		const lines = clippedSeriesLine(crossing, true, style, band);
		expect(lines).toHaveLength(1);
		const path = lines[0];
		expect(path?.kind).toBe('path');
		const ys = [...(path?.kind === 'path' ? path.d : '').matchAll(/,(-?[\d.]+)/gu)].map((m) =>
			Number(m[1]),
		);
		expect(Math.min(...ys)).toBeGreaterThanOrEqual(band.top - 1e-6);
		expect(Math.max(...ys)).toBeLessThanOrEqual(band.bottom + 1e-6);
	});
});

describe('stacked lines and areas with c:min / c:max', () => {
	const element = (chartData: PptxChartData): ChartPptxElement =>
		({
			id: 'c',
			type: 'chart',
			x: 0,
			y: 0,
			width: 400,
			height: 300,
			chartData,
		}) as ChartPptxElement;
	const chart = (chartType: 'line' | 'area', min: number, max: number): PptxChartData => ({
		chartType,
		categories: ['A', 'B', 'C'],
		series: [
			{ name: 'S1', values: [40, 60, 50] },
			{ name: 'S2', values: [50, 70, 60] },
		],
		grouping: 'stacked',
		style: { hasDataLabels: true },
		axes: [{ axisType: 'valAx', axPos: 'l', majorGridlines: true, min, max }],
	});
	const plotYs = (primitives: ReadonlyArray<SvgPrimitive>): number[] =>
		primitives.flatMap((p) => {
			if (p.kind === 'polyline') {
				return p.points.split(' ').map((pair) => Number(pair.split(',')[1]));
			}
			if (p.kind === 'path') {
				return [...p.d.matchAll(/,(-?[\d.]+)/gu)].map((m) => Number(m[1]));
			}
			if (p.kind === 'circle') {
				return [p.cy];
			}
			return [];
		});

	for (const chartType of ['line', 'area'] as const) {
		it(`keeps a ${chartType} chart inside the plot when c:max is below the totals`, () => {
			const vm = buildChartViewModel(element(chart(chartType, 0, 100)));
			const top = Math.min(...vm.gridlines.map((g) => g.y1));
			const bottom = Math.max(...vm.gridlines.map((g) => g.y1));
			const ys = plotYs(vm.primitives);
			expect(ys.length).toBeGreaterThan(0);
			expect(Math.min(...ys)).toBeGreaterThanOrEqual(top - 0.01);
			expect(Math.max(...ys)).toBeLessThanOrEqual(bottom + 0.01);
			for (const label of vm.dataLabels) {
				expect(label.y).toBeGreaterThanOrEqual(top - 0.01);
				expect(label.y).toBeLessThanOrEqual(bottom + 0.01);
			}
		});

		it(`keeps a ${chartType} chart inside the plot when c:min is above zero`, () => {
			const vm = buildChartViewModel(element(chart(chartType, 50, 150)));
			const top = Math.min(...vm.gridlines.map((g) => g.y1));
			const bottom = Math.max(...vm.gridlines.map((g) => g.y1));
			const ys = plotYs(vm.primitives);
			expect(ys.length).toBeGreaterThan(0);
			expect(Math.min(...ys)).toBeGreaterThanOrEqual(top - 0.01);
			expect(Math.max(...ys)).toBeLessThanOrEqual(bottom + 0.01);
		});
	}
});
