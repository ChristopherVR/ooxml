/**
 * Combination charts, secondary axes and series-level data labels of a NEW chart, through the
 * public SDK, a save and a reload: the XML and the reloaded model are both checked, and the
 * combinations the shared core cannot write are reported.
 */
import { describe, expect, it } from 'vitest';

import { setChartSeriesChartType } from '../../core/builders/sdk/chart-operations';
import { withChartDataLabelOptions } from '../../core/builders/sdk/chart-format-options-operations';
import {
	chartSupportsSecondaryAxis,
	isChartSeriesOnSecondaryAxis,
	withChartSeriesOnSecondaryAxis,
} from '../../core/builders/sdk/chart-series-axis-operations';
import { withChartElementFormat } from '../../core/builders/sdk/chart-element-format-operations';
import { firstElement, newChartDeck, type NewChartPath } from '../helpers/new-chart-deck';

const groups = (xml: string) =>
	[...xml.matchAll(/<c:(\w+Chart)>/gu)].map((match) => match[1]).filter((tag) => tag !== 'Chart');

describe.each(['addChart', 'builder'] as NewChartPath[])(
	'combination charts through %s',
	(path) => {
		it('writes the column series around a line series as one column group', async () => {
			const { element, save } = await newChartDeck('bar', {}, path);
			setChartSeriesChartType(element, 1, 'line');
			expect(element.chartData!.chartType).toBe('combo');
			const { xml, chart, warnings } = await save();
			// PowerPoint draws two column groups on the same axes over each other.
			expect(groups(xml)).toStrictEqual(['barChart', 'lineChart']);
			expect(chart.chartType).toBe('combo');
			// Reloaded group by group; every series keeps its own c:order.
			expect(chart.series.map((series) => [series.name, series.seriesChartType])).toStrictEqual([
				['Revenue', 'bar'],
				['Cost', 'bar'],
				['Margin', 'line'],
			]);
			expect(firstElement(xml, 'c:lineChart')).toContain('<c:order val="1"></c:order>');
			expect(firstElement(xml, 'c:lineChart')).toContain('<c:smooth val="0"></c:smooth>');
			expect(warnings).toStrictEqual([]);
		});

		it('writes three families and a line chart whose first series becomes columns', async () => {
			const mixed = await newChartDeck('bar', {}, path);
			setChartSeriesChartType(mixed.element, 1, 'area');
			setChartSeriesChartType(mixed.element, 2, 'line');
			const saved = await mixed.save();
			expect(saved.chart.series.map((series) => series.seriesChartType)).toStrictEqual([
				'bar',
				'area',
				'line',
			]);
			const line = await newChartDeck('line', {}, path);
			setChartSeriesChartType(line.element, 0, 'bar');
			const first = await line.save();
			expect(groups(first.xml)).toStrictEqual(['barChart', 'lineChart']);
			expect(first.chart.series.map((series) => series.seriesChartType)).toStrictEqual([
				'bar',
				'line',
				'line',
			]);
		});
	},
);

describe('combinations the shared core cannot write', () => {
	it.each(['scatter', 'pie'] as const)(
		'keeps the legacy generator for a column and %s mix, and reports it',
		async (type) => {
			const { element, save } = await newChartDeck('bar');
			setChartSeriesChartType(element, 1, type);
			const { chart, warnings } = await save();
			expect(warnings.map((warning) => warning.code)).toStrictEqual([
				'CHART_WRITTEN_BY_LEGACY_GENERATOR',
			]);
			expect(warnings[0]!.message).toContain('Series combo types need');
			// The legacy generator writes the column chart it always wrote.
			expect(chart.chartType).toBe('bar');
		},
	);
});

describe('secondary value axes', () => {
	it('lets the SDK move a series of a new chart onto a secondary axis, and saves it', async () => {
		const { element, save } = await newChartDeck('bar');
		expect(element.chartData!.axes).toBeUndefined();
		expect(chartSupportsSecondaryAxis(element.chartData!)).toBe(true);
		element.chartData = withChartSeriesOnSecondaryAxis(element.chartData!, 1, true);
		expect(isChartSeriesOnSecondaryAxis(element.chartData, 1)).toBe(true);
		expect(element.chartData.axes!.map((axis) => axis.axisId)).toStrictEqual([
			111111111, 222222222, 222222223, 222222224,
		]);
		const { xml, chart, warnings } = await save();
		expect(warnings).toStrictEqual([]);
		expect(xml.match(/<c:valAx>/gu)).toHaveLength(2);
		const secondary = chart.axes!.find((axis) => axis.axisId === 222222223)!;
		expect(secondary).toMatchObject({ axisType: 'valAx', axPos: 'r', crosses: 'max' });
		expect(chart.axes!.find((axis) => axis.axisId === 222222224)).toMatchObject({ deleted: true });
		// The two primary column series form one group; the moved series follows in its own.
		expect(xml.match(/<c:barChart>/gu)).toHaveLength(2);
		expect(chart.series.map((series) => series.name)).toStrictEqual(['Revenue', 'Cost', 'Margin']);
		expect(chart.series[2]!.axisId).toBe(222222223);
		expect(isChartSeriesOnSecondaryAxis(chart, 2)).toBe(true);
		expect(isChartSeriesOnSecondaryAxis(chart, 0)).toBe(false);
	});

	it('puts the line series of a combination chart on the secondary axis', async () => {
		const { element, save } = await newChartDeck('bar');
		setChartSeriesChartType(element, 2, 'line');
		element.chartData = withChartSeriesOnSecondaryAxis(element.chartData!, 2, true);
		const { xml, chart } = await save();
		expect(firstElement(xml, 'c:lineChart')).toContain('<c:axId val="222222223"></c:axId>');
		expect(chart.series[2]).toMatchObject({ seriesChartType: 'line', axisId: 222222223 });
	});

	it('writes a title set on the secondary axis entry and keeps primary axis formats in place', async () => {
		const { element, save } = await newChartDeck('bar');
		element.chartData = withChartElementFormat(
			element.chartData!,
			{ kind: 'axis', axisIndex: 1 },
			{ line: { kind: 'solid', color: '#00AA00', width: 1 } },
		);
		element.chartData = withChartSeriesOnSecondaryAxis(element.chartData, 1, true);
		element.chartData.axes![2]!.titleText = 'Margin %';
		const { xml, chart } = await save();
		expect(chart.axes!.find((axis) => axis.axisId === 222222223)!.titleText).toBe('Margin %');
		const valueAxes = [...xml.matchAll(/<c:valAx>.*?<\/c:valAx>/gsu)].map((match) => match[0]);
		expect(valueAxes[0]).toContain('00AA00');
		expect(valueAxes[1]).not.toContain('00AA00');
	});

	it('removes the pair again when the series goes back', async () => {
		const { element, save } = await newChartDeck('line');
		element.chartData = withChartSeriesOnSecondaryAxis(element.chartData!, 1, true);
		element.chartData = withChartSeriesOnSecondaryAxis(element.chartData, 1, false);
		const { xml, chart } = await save();
		expect(xml.match(/<c:valAx>/gu)).toHaveLength(1);
		expect(chart.axes!.filter((axis) => axis.axisType === 'valAx')).toHaveLength(1);
	});

	it.each([
		['the first series', {}, 0],
		['horizontal bars', { barDirection: 'bar' as const }, 1],
	])(
		'reports %s on a secondary axis, which the legacy generator drops',
		async (_, input, index) => {
			const { element, save } = await newChartDeck('bar', input);
			element.chartData = withChartSeriesOnSecondaryAxis(element.chartData!, index, true);
			expect(element.chartData.series[index]!.axisId).toBe(222222223);
			const { chart, warnings } = await save();
			expect(warnings.map((warning) => warning.code)).toStrictEqual([
				'CHART_WRITTEN_BY_LEGACY_GENERATOR',
			]);
			expect(warnings[0]!.message).toContain('Secondary-axis editing needs');
			expect(chart.axes!.filter((axis) => axis.axisType === 'valAx')).toHaveLength(1);
		},
	);

	it('leaves a 3-D chart, which PowerPoint gives no secondary axis, as it is', async () => {
		const { element } = await newChartDeck('bar3D');
		expect(chartSupportsSecondaryAxis(element.chartData!)).toBe(false);
		expect(withChartSeriesOnSecondaryAxis(element.chartData!, 1, true)).toBe(element.chartData);
	});

	it('leaves a pie chart, which has no axes, as it is', async () => {
		const { element } = await newChartDeck('pie');
		expect(chartSupportsSecondaryAxis(element.chartData!)).toBe(false);
		expect(withChartSeriesOnSecondaryAxis(element.chartData!, 1, true)).toBe(element.chartData);
	});
});

describe('series data labels', () => {
	it("writes one series' own labels with content, separator and number format", async () => {
		const { element, save } = await newChartDeck('bar');
		element.chartData = withChartDataLabelOptions(
			element.chartData!,
			{ showValue: true, showCategory: true },
			2,
		);
		Object.assign(element.chartData.series[2]!.dataLabelOptions!, {
			separator: '; ',
			numberFormat: '0.0',
		});
		const { xml, chart, warnings } = await save();
		expect(warnings).toStrictEqual([]);
		const third = xml.slice(xml.lastIndexOf('<c:ser>'));
		expect(third).toContain('<c:separator>; </c:separator>');
		expect(third).toContain('<c:numFmt formatCode="0.0" sourceLinked="0"></c:numFmt>');
		expect(chart.series[2]!.dataLabelOptions).toMatchObject({
			showValue: true,
			showCategory: true,
			separator: '; ',
			numberFormat: '0.0',
		});
		// The other series do not show labels: the chart-wide flags are written off.
		expect(chart.style?.dataLabels?.showValue).toBe(false);
		expect(chart.series[0]!.dataLabelOptions?.showValue).not.toBe(true);
	});

	it('writes percentages on a pie series', async () => {
		const { element, save } = await newChartDeck('pie');
		element.chartData = withChartDataLabelOptions(element.chartData!, { showPercent: true }, 0);
		const { chart } = await save();
		expect(chart.series[0]!.dataLabelOptions?.showPercent).toBe(true);
	});
});
