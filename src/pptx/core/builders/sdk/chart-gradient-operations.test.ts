import { describe, expect, it } from 'vitest';

import type { PptxChartType } from '../../types/chart';
import type { ChartPptxElement } from '../../types/elements';
import { setChartDataPointStyle } from './chart-formatting-operations';
import type { ChartGradientInput } from './chart-gradient-input';
import { toChartGradientFill } from './chart-gradient-input';
import { setChartDataPointGradient, setChartSeriesGradient } from './chart-gradient-operations';
import {
	setChartDataPointFill,
	setChartDataPointExplosion,
	setChartSeriesChartType,
	setChartSeriesColor,
} from './chart-operations';
import { ChartBuilder } from './ChartBuilder';
import { createChartElement } from './ElementFactory';

const gradient: ChartGradientInput = {
	stops: [
		{ color: '#1e3a8a', position: 100 },
		{ color: '60A5FA', position: 0 },
	],
};

const expected = {
	type: 'linear',
	angle: 90,
	stops: [
		{ color: '#60A5FA', position: 0 },
		{ color: '#1E3A8A', position: 100 },
	],
};

function makeChart(chartType: PptxChartType = 'bar'): ChartPptxElement {
	return createChartElement(chartType, {
		series: [{ name: 'Revenue', values: [100, 200, 300], color: '#4472C4' }],
		categories: ['Q1', 'Q2', 'Q3'],
	});
}

describe('toChartGradientFill', () => {
	it('normalises a linear input the way a reload parses it back', () => {
		expect(toChartGradientFill(gradient)).toStrictEqual(expected);
		expect(toChartGradientFill({ ...gradient, angle: -90 }).angle).toBe(270);
	});

	it('defaults a radial focal point to the centre and keeps stop opacity below 1', () => {
		expect(
			toChartGradientFill({
				type: 'radial',
				stops: [
					{ color: '#FFFFFF', position: 0, opacity: 0.4 },
					{ color: '#000000', position: 100, opacity: 1 },
				],
			}),
		).toStrictEqual({
			type: 'radial',
			focalPoint: { x: 0.5, y: 0.5 },
			stops: [
				{ color: '#FFFFFF', position: 0, opacity: 0.4 },
				{ color: '#000000', position: 100 },
			],
		});
	});

	it('rejects invalid gradients', () => {
		expect(() => toChartGradientFill({ stops: [{ color: '#FFFFFF', position: 0 }] })).toThrow(
			RangeError,
		);
		expect(() =>
			toChartGradientFill({
				stops: [
					{ color: 'red', position: 0 },
					{ color: '#000000', position: 100 },
				],
			}),
		).toThrow(/6-digit hex/u);
		expect(() =>
			toChartGradientFill({
				stops: [
					{ color: '#FFFFFF', position: -1 },
					{ color: '#000000', position: 100 },
				],
			}),
		).toThrow(RangeError);
		expect(() =>
			toChartGradientFill({
				stops: [
					{ color: '#FFFFFF', position: 0, opacity: 2 },
					{ color: '#000000', position: 100 },
				],
			}),
		).toThrow(RangeError);
		expect(() => toChartGradientFill({ ...gradient, angle: Number.NaN })).toThrow(RangeError);
		expect(() =>
			toChartGradientFill({ type: 'radial', focalPoint: { x: 2, y: 0 }, stops: gradient.stops }),
		).toThrow(RangeError);
	});
});

describe('createChartElement / ChartBuilder with a series gradient', () => {
	it('maps ChartSeriesInput.gradientFill onto the series model', () => {
		const chart = createChartElement('pie', {
			series: [{ name: 'S', values: [1, 2], gradientFill: gradient }],
			categories: ['A', 'B'],
		});
		expect(chart.chartData!.series[0].gradientFill).toStrictEqual(expected);
	});

	it('applies .gradient() to the most recently added series', () => {
		const chart = ChartBuilder.create('bar')
			.categories(['A', 'B'])
			.addSeries('First', [1, 2], '#FF0000')
			.addSeries('Second', [3, 4])
			.gradient(gradient)
			.build();
		expect(chart.chartData!.series[0].gradientFill).toBeUndefined();
		expect(chart.chartData!.series[1].gradientFill).toStrictEqual(expected);
	});

	it('throws for .gradient() before any series and for line-drawn chart types', () => {
		expect(() => ChartBuilder.create('bar').gradient(gradient)).toThrow(/addSeries/u);
		for (const type of ['line', 'line3D', 'scatter', 'radar', 'stock', 'waterfall'] as const) {
			expect(() =>
				createChartElement(type, {
					series: [{ name: 'S', values: [1, 2], gradientFill: gradient }],
					categories: ['A', 'B'],
				}),
			).toThrow(/area-filled chart types/u);
		}
	});
});

describe('setChartSeriesGradient', () => {
	it('sets and clears a series gradient, keeping the solid colour as the fallback', () => {
		const chart = makeChart();
		setChartSeriesGradient(chart, 0, gradient);
		expect(chart.chartData!.series[0].gradientFill).toStrictEqual(expected);
		setChartSeriesGradient(chart, 0, null);
		expect(chart.chartData!.series[0].gradientFill).toBeUndefined();
		expect(chart.chartData!.series[0].color).toBe('#4472C4');
	});

	it('is cleared by setChartSeriesColor: the last call wins', () => {
		const chart = makeChart();
		setChartSeriesGradient(chart, 0, gradient);
		setChartSeriesColor(chart, 0, '#00B050');
		expect(chart.chartData!.series[0].gradientFill).toBeUndefined();
		expect(chart.chartData!.series[0].color).toBe('#00B050');

		setChartSeriesGradient(chart, 0, gradient);
		setChartSeriesColor(chart, 0, null);
		expect(chart.chartData!.series[0].gradientFill).toBeUndefined();
	});

	it('throws for an out-of-range series and for a line-drawn combo series', () => {
		const chart = makeChart();
		expect(() => setChartSeriesGradient(chart, 3, gradient)).toThrow(RangeError);
		expect(() => setChartSeriesGradient(makeChart('radar'), 0, gradient)).toThrow(
			/drawn as "radar"/u,
		);
		setChartSeriesChartType(chart, 0, 'line');
		expect(() => setChartSeriesGradient(chart, 0, gradient)).toThrow(/drawn as "line"/u);
		expect(() => setChartSeriesGradient(chart, 0, null)).not.toThrow();
	});
});

describe('setChartDataPointGradient', () => {
	it('creates a c:dPt override and replaces a per-point solid fill', () => {
		const chart = makeChart();
		setChartDataPointStyle(chart, 0, 2, { fillColor: '#FF0000', strokeWidth: 2 });
		setChartDataPointGradient(chart, 0, 2, gradient);
		expect(chart.chartData!.series[0].dataPoints).toStrictEqual([
			{ idx: 2, spPr: { strokeWidth: 2 }, gradientFill: expected },
		]);
	});

	it('is cleared by a later solid fill on the same point', () => {
		const chart = makeChart();
		setChartDataPointGradient(chart, 0, 1, gradient);
		setChartDataPointFill(chart, 0, 1, '#00B050');
		expect(chart.chartData!.series[0].dataPoints).toStrictEqual([
			{ idx: 1, spPr: { fillColor: '#00B050' } },
		]);
		setChartDataPointGradient(chart, 0, 1, gradient);
		setChartDataPointStyle(chart, 0, 1, null);
		expect(chart.chartData!.series[0].dataPoints).toStrictEqual([]);
	});

	it('drops the empty override when cleared, keeping one with other formatting', () => {
		const chart = makeChart('pie');
		setChartDataPointGradient(chart, 0, 0, gradient);
		setChartDataPointGradient(chart, 0, 1, gradient);
		setChartDataPointExplosion(chart, 0, 1, 20);
		setChartDataPointGradient(chart, 0, 0, null);
		setChartDataPointGradient(chart, 0, 1, null);
		expect(chart.chartData!.series[0].dataPoints).toStrictEqual([{ idx: 1, explosion: 20 }]);
	});

	it('validates the point index and chart type', () => {
		expect(() => setChartDataPointGradient(makeChart(), 0, -1, gradient)).toThrow(RangeError);
		expect(() => setChartDataPointGradient(makeChart(), 0, 1.5, gradient)).toThrow(RangeError);
		expect(() => setChartDataPointGradient(makeChart('scatter'), 0, 0, gradient)).toThrow(
			/area-filled chart types/u,
		);
	});
});
