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
		expect(xml).toMatch(/<c:gapWidth val="150"\s*(?:\/>|><\/c:gapWidth>)/u);
		expect(xml).not.toContain('<c:overlap');
		expect(chart.barGapWidth).toBe(150);
	});

	it('starts a doughnut at the documented hole size 50, and a pie with no gap or hole', async () => {
		const doughnut = await newChartDeck('doughnut', {}, path);
		expect(doughnut.element.chartData!.doughnutHoleSize).toBe(50);
		const saved = await doughnut.save();
		expect(saved.xml).toMatch(/<c:holeSize val="50"\s*(?:\/>|><\/c:holeSize>)/u);
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
			expect(xml).toMatch(/<c:overlap val="100"\s*(?:\/>|><\/c:overlap>)/u);
			expect(chart.barOverlap).toBe(100);
		}
	});

	it('writes a straight line unless the series asks for smoothing', async () => {
		const { element, save } = await newChartDeck('line', {}, path);
		element.chartData!.series[1]!.smooth = true;
		const { xml, chart } = await save();
		expect(xml.match(/<c:smooth val="0"\/>/gu)).toHaveLength(2);
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
			expect(xml).toMatch(/<c:grouping val="standard"\s*(?:\/>|><\/c:grouping>)/u);
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
		expect(xml).toMatch(/<c:overlap val="0"\s*(?:\/>|><\/c:overlap>)/u);
		expect(chart.barOverlap).toBe(0);
	});

	it('writes and reloads the bounds of each range', async () => {
		const bar = await newChartDeck('bar', { gapWidth: 0, overlap: -100 });
		const low = await bar.save();
		expect(low.xml).toContain('<c:gapWidth val="0"/><c:overlap val="-100"/>');
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
		expect(pie.xml).toMatch(/<c:firstSliceAng val="0"\s*(?:\/>|><\/c:firstSliceAng>)/u);
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

	it('writes bar series as not inverted, so PowerPoint keeps the colour of negative bars', async () => {
		const series = (xml: string): string[] => xml.match(/<c:ser>.*?<\/c:ser>/gsu) ?? [];
		const bar = await (await newChartDeck('bar')).save();
		expect(series(bar.xml)).toHaveLength(3);
		for (const item of series(bar.xml)) expect(item).toContain('<c:invertIfNegative val="0"/>');
		expect(bar.chart.series.map((item) => item.invertIfNegative)).toStrictEqual([
			false,
			false,
			false,
		]);

		// A series' own setting is kept, and a line series has no such element.
		const combo = await newChartDeck('line', {
			series: [
				{ name: 'L', values: [1, -2, 3] },
				{ name: 'B', values: [-1, 2, -3], chartType: 'bar' },
				{ name: 'I', values: [-1, 2, -3], chartType: 'bar' },
			],
		});
		combo.element.chartData!.series[2]!.invertIfNegative = true;
		const saved = await combo.save();
		expect(saved.warnings).toStrictEqual([]);
		const byName = (name: string) =>
			series(saved.xml).find((item) => item.includes(`<c:v>${name}</c:v>`))!;
		expect(byName('L')).not.toContain('invertIfNegative');
		expect(byName('B')).toContain('<c:invertIfNegative val="0"/>');
		expect(byName('I')).toContain('<c:invertIfNegative val="1"/>');
	});

	it('states the 1900 date system and square corners, which PowerPoint reads otherwise when absent', async () => {
		const plain = await (await newChartDeck('line')).save();
		expect(plain.xml).toMatch(
			/<c:chartSpace[^>]*><c:date1904 val="0"\/><c:roundedCorners val="0"\/>/u,
		);
		expect(plain.chart.roundedCorners).toBe(false);

		// The model's own values are kept.
		const own = await newChartDeck('line', { chartArea: { roundedCorners: true } });
		own.element.chartData!.date1904 = true;
		const saved = await own.save();
		expect(saved.warnings).toStrictEqual([]);
		expect(saved.xml).toMatch(
			/<c:chartSpace[^>]*><c:date1904 val="1"\/><c:roundedCorners val="1"\/>/u,
		);
		expect(saved.chart.roundedCorners).toBe(true);
		expect(saved.chart.date1904).toBe(true);
	});

	it('keeps the legacy generator from writing clustered into a line chart', () => {
		const element = createChartElement('line', { ...NEW_CHART_INPUT, grouping: 'clustered' });
		const xml = new PptxRuntimeDependencyFactory()
			.createBuilder()
			.build(buildChartSpaceXml(element.chartData!)) as string;
		expect(xml).toMatch(/<c:grouping val="standard"\s*(?:\/>|><\/c:grouping>)/u);
	});
});
