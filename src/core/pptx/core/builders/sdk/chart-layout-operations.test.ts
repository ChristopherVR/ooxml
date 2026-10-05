import { describe, expect, it } from 'vitest';

import type { PptxChartType } from '../../types/chart';
import type { ChartPptxElement } from '../../types/elements';
import { ChartBuilder } from './ChartBuilder';
import { setChartAreaFormat, setChartGroupOptions } from './chart-layout-operations';
import { setChartAxis, setChartSeriesChartType } from './chart-operations';
import { createChartElement } from './ElementFactory';

function makeChart(chartType: PptxChartType = 'bar'): ChartPptxElement {
	return createChartElement(chartType, {
		categories: ['A', 'B'],
		series: [{ name: 'S', values: [1, 2] }],
	});
}

describe('setChartGroupOptions', () => {
	it('sets and clears values on the model', () => {
		const chart = makeChart();
		setChartGroupOptions(chart, { gapWidth: 35, overlap: -100 });
		expect(chart.chartData).toMatchObject({ barGapWidth: 35, barOverlap: -100 });
		setChartGroupOptions(chart, { overlap: null });
		expect(chart.chartData!.barOverlap).toBeUndefined();
		expect(chart.chartData!.barGapWidth).toBe(35);
	});

	it('validates against the schema ranges', () => {
		const bar = makeChart();
		expect(() => setChartGroupOptions(bar, { gapWidth: 501 })).toThrow(RangeError);
		expect(() => setChartGroupOptions(bar, { gapWidth: 12.5 })).toThrow(/ST_GapAmount/u);
		expect(() => setChartGroupOptions(bar, { overlap: -101 })).toThrow(/ST_Overlap/u);
		const doughnut = makeChart('doughnut');
		expect(() => setChartGroupOptions(doughnut, { holeSize: 0 })).toThrow(/1 to 90/u);
		expect(() => setChartGroupOptions(doughnut, { firstSliceAngle: 361 })).toThrow(
			/ST_FirstSliceAng/u,
		);
		setChartGroupOptions(doughnut, { holeSize: 90, firstSliceAngle: 360 });
		expect(doughnut.chartData).toMatchObject({ doughnutHoleSize: 90, firstSliceAngle: 360 });
	});

	it('rejects options the chart type has no element for', () => {
		expect(() => setChartGroupOptions(makeChart(), { holeSize: 40 })).toThrow(
			/holeSize applies to doughnut charts; this chart is drawn as "bar"/u,
		);
		expect(() => setChartGroupOptions(makeChart('bar3D'), { overlap: 10 })).toThrow(
			/overlap applies to bar charts/u,
		);
		expect(() => setChartGroupOptions(makeChart('line'), { gapWidth: 10 })).toThrow(Error);
		expect(() => setChartGroupOptions(makeChart('pie3D'), { firstSliceAngle: 10 })).toThrow(Error);
	});

	it('accepts an option a combo chart draws through one of its series', () => {
		const chart = createChartElement('line', {
			categories: ['A'],
			series: [
				{ name: 'L', values: [1] },
				{ name: 'B', values: [2] },
			],
		});
		setChartSeriesChartType(chart, 1, 'bar');
		setChartGroupOptions(chart, { gapWidth: 60 });
		expect(chart.chartData!.barGapWidth).toBe(60);
	});

	it('keeps the of-pie options gap width in step', () => {
		const chart = makeChart('ofPie');
		chart.chartData!.ofPieOptions = { ofPieType: 'bar', gapWidth: 100 };
		setChartGroupOptions(chart, { gapWidth: 40 });
		expect(chart.chartData!.ofPieOptions.gapWidth).toBe(40);
	});
});

describe('setChartAreaFormat', () => {
	it('sets fill, gradient, border and rounded corners, and clears them with null', () => {
		const chart = makeChart();
		setChartAreaFormat(chart, 'chart', { fill: 'ffffff', border: 'none', roundedCorners: false });
		expect(chart.chartData!.style).toMatchObject({
			chartAreaFill: '#FFFFFF',
			chartAreaBorder: 'none',
		});
		expect(chart.chartData!.roundedCorners).toBe(false);

		setChartAreaFormat(chart, 'chart', {
			fill: {
				stops: [
					{ color: '#FFFFFF', position: 0 },
					{ color: '#000000', position: 100 },
				],
			},
		});
		expect(chart.chartData!.style!.chartAreaFill).toBeUndefined();
		expect(chart.chartData!.style!.chartAreaGradient?.angle).toBe(90);

		setChartAreaFormat(chart, 'chart', null);
		expect(chart.chartData!.style!.chartAreaGradient).toBeUndefined();
		expect(chart.chartData!.style!.chartAreaBorder).toBeUndefined();
		expect(chart.chartData!.roundedCorners).toBeUndefined();
	});

	it('formats the plot area and rejects bad input', () => {
		const chart = makeChart();
		setChartAreaFormat(chart, 'plot', { fill: 'none', border: '#D9D9D9' });
		expect(chart.chartData!.style).toMatchObject({
			plotAreaFill: 'none',
			plotAreaBorder: '#D9D9D9',
		});
		expect(() => setChartAreaFormat(chart, 'plot', { roundedCorners: false })).toThrow(
			/chart area only/u,
		);
		expect(() => setChartAreaFormat(chart, 'chart', { fill: 'red' })).toThrow(/hex colour/u);
		expect(() => setChartAreaFormat(chart, 'chart', { fill: { stops: [] } })).toThrow(RangeError);
	});
});

describe('ChartInput and ChartBuilder layout fields', () => {
	it('applies ChartInput fields through createChartElement', () => {
		const chart = createChartElement('bar', {
			categories: ['A'],
			series: [{ name: 'S', values: [1] }],
			gapWidth: 45,
			chartArea: { fill: 'none', roundedCorners: false },
			axes: { catAx: { visible: false }, valAx: { max: 1 } },
		});
		expect(chart.chartData).toMatchObject({ barGapWidth: 45, roundedCorners: false });
		expect(chart.chartData!.axes).toStrictEqual([
			{ axisType: 'catAx', deleted: true },
			{ axisType: 'valAx', max: 1 },
		]);
		expect(() =>
			createChartElement('pie', {
				categories: ['A'],
				series: [{ name: 'S', values: [1] }],
				holeSize: 50,
			}),
		).toThrow(/doughnut/u);
	});

	it('builds the same through ChartBuilder', () => {
		const chart = ChartBuilder.create('doughnut')
			.categories(['A', 'B'])
			.addSeries('S', [1, 2])
			.firstSliceAngle(90)
			.holeSize(25)
			.chartArea({ fill: 'none', border: 'none', roundedCorners: false })
			.plotArea({ fill: 'none' })
			.build();
		expect(chart.chartData).toMatchObject({ firstSliceAngle: 90, doughnutHoleSize: 25 });
		const bar = ChartBuilder.create('bar')
			.categories(['A'])
			.addSeries('S', [1])
			.gapWidth(30)
			.overlap(-4)
			.axis('valAx', { visible: false })
			.axis('valAx', { min: 0 })
			.build();
		expect(bar.chartData).toMatchObject({ barGapWidth: 30, barOverlap: -4 });
		expect(bar.chartData!.axes).toStrictEqual([{ axisType: 'valAx', deleted: true, min: 0 }]);
		expect(() =>
			ChartBuilder.create('bar').categories(['A']).addSeries('S', [1]).gapWidth(600).build(),
		).toThrow(RangeError);
	});

	it('shows or hides an axis through setChartAxis', () => {
		const chart = makeChart();
		setChartAxis(chart, 'valAx', { visible: false });
		expect(chart.chartData!.axes![0]!.deleted).toBe(true);
		setChartAxis(chart, 'valAx', { visible: true });
		expect(chart.chartData!.axes![0]!.deleted).toBe(false);
	});
});
