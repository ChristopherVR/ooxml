/**
 * Documented defaults and boundaries of a NEW chart's group options, and the grouping rule for
 * line and area charts, through the public SDK, a save and a reload, on both creation paths.
 */
import { describe, expect, it } from 'vitest';

import { PptxRuntimeDependencyFactory } from '../../core/core/factories/PptxRuntimeDependencyFactory';
import { createChartElement } from '../../core/builders/sdk/ElementFactory';
import { buildChartSpaceXml } from '../../core/utils/chart-xml-generator';
import { NEW_CHART_INPUT, newChartDeck, type NewChartPath } from '../helpers/new-chart-deck';

const PATHS: NewChartPath[] = ['addChart', 'builder'];

describe.each(PATHS)('new chart defaults through %s', (path) => {
	it('starts a bar chart at the documented gap width 150 with no overlap', async () => {
		const { element, save } = await newChartDeck('bar', {}, path);
		expect(element.chartData!.barGapWidth).toBe(150);
		expect(element.chartData!.barOverlap).toBeUndefined();
		const { xml, chart } = await save();
		expect(xml).toContain('<c:gapWidth val="150"></c:gapWidth>');
		expect(xml).not.toContain('<c:overlap');
		expect(chart.barGapWidth).toBe(150);
	});

	it('starts a doughnut at the documented hole size 50, and a pie with no gap or hole', async () => {
		const doughnut = await newChartDeck('doughnut', {}, path);
		expect(doughnut.element.chartData!.doughnutHoleSize).toBe(50);
		const saved = await doughnut.save();
		expect(saved.xml).toContain('<c:holeSize val="50"></c:holeSize>');
		expect(saved.chart.doughnutHoleSize).toBe(50);
		const pie = await newChartDeck('pie', {}, path);
		expect(pie.element.chartData!.barGapWidth).toBeUndefined();
		expect(pie.element.chartData!.doughnutHoleSize).toBeUndefined();
		expect((await pie.save()).xml).not.toMatch(/<c:(gapWidth|holeSize)/u);
	});

	it('overlaps stacked bars fully unless the input sets an overlap', async () => {
		for (const grouping of ['stacked', 'percentStacked'] as const) {
			const { element, save } = await newChartDeck('bar', { grouping }, path);
			expect(element.chartData!.barOverlap).toBeUndefined();
			const { xml, chart } = await save();
			expect(xml).toContain('<c:overlap val="100"></c:overlap>');
			expect(chart.barOverlap).toBe(100);
		}
	});

	it('writes a straight line unless the series asks for smoothing', async () => {
		const { element, save } = await newChartDeck('line', {}, path);
		element.chartData!.series[1]!.smooth = true;
		const { xml, chart } = await save();
		expect(xml.match(/<c:smooth val="0"><\/c:smooth>/gu)).toHaveLength(2);
		expect(chart.series.map((series) => series.smooth)).toStrictEqual([false, true, false]);
	});

	it('writes scatter smoothing as the scatter style draws it', async () => {
		const straight = await (await newChartDeck('scatter', {}, path)).save();
		expect(straight.chart.series.map((series) => series.smooth)).toStrictEqual([
			false,
			false,
			false,
		]);
		const curved = await newChartDeck('scatter', {}, path);
		curved.element.chartData!.scatterStyle = 'smoothMarker';
		curved.element.chartData!.series[0]!.smooth = false;
		const saved = await curved.save();
		expect(saved.chart.series.map((series) => series.smooth)).toStrictEqual([false, true, true]);
	});

	it('writes a line or area chart asked for clustered as standard', async () => {
		for (const type of ['line', 'area'] as const) {
			const { save } = await newChartDeck(type, { grouping: 'clustered' }, path);
			const { xml, chart } = await save();
			expect(xml).toContain('<c:grouping val="standard"></c:grouping>');
			// The model names the lowest grouping of every family `clustered`.
			expect([chart.grouping, chart.groupingStandard]).toStrictEqual(['clustered', true]);
		}
	});
});

describe('new chart group option boundaries', () => {
	it('keeps an explicit overlap on stacked bars, zero included', async () => {
		const { xml, chart } = await (
			await newChartDeck('bar', { grouping: 'stacked', overlap: 0 })
		).save();
		expect(xml).toContain('<c:overlap val="0"></c:overlap>');
		expect(chart.barOverlap).toBe(0);
	});

	it('writes and reloads the bounds of each range', async () => {
		const bar = await newChartDeck('bar', { gapWidth: 0, overlap: -100 });
		const low = await bar.save();
		expect(low.xml).toContain(
			'<c:gapWidth val="0"></c:gapWidth><c:overlap val="-100"></c:overlap>',
		);
		expect([low.chart.barGapWidth, low.chart.barOverlap]).toStrictEqual([0, -100]);
		const high = await (await newChartDeck('bar', { gapWidth: 500, overlap: 100 })).save();
		expect([high.chart.barGapWidth, high.chart.barOverlap]).toStrictEqual([500, 100]);
		const doughnut = await (
			await newChartDeck('doughnut', { holeSize: 1, firstSliceAngle: 360 })
		).save();
		expect([doughnut.chart.doughnutHoleSize, doughnut.chart.firstSliceAngle]).toStrictEqual([
			1, 360,
		]);
		const pie = await (await newChartDeck('pie', { firstSliceAngle: 0 })).save();
		expect(pie.xml).toContain('<c:firstSliceAng val="0"></c:firstSliceAng>');
		expect(pie.chart.firstSliceAngle).toBe(0);
	});

	it('refuses fractional or out-of-range values and options the type does not have', () => {
		for (const input of [
			{ gapWidth: 12.5 },
			{ gapWidth: 501 },
			{ overlap: -101 },
			{ holeSize: 0 },
			{ holeSize: 91 },
			{ firstSliceAngle: 361 },
		])
			expect(() => createChartElement('doughnut', { ...NEW_CHART_INPUT, ...input })).toThrow();
		expect(() => createChartElement('line', { ...NEW_CHART_INPUT, gapWidth: 80 })).toThrow(
			/applies to/u,
		);
	});

	it('keeps the legacy generator from writing clustered into a line chart', () => {
		const element = createChartElement('line', { ...NEW_CHART_INPUT, grouping: 'clustered' });
		const xml = new PptxRuntimeDependencyFactory()
			.createBuilder()
			.build(buildChartSpaceXml(element.chartData!)) as string;
		expect(xml).toContain('<c:grouping val="standard"></c:grouping>');
	});
});
