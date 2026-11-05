import { describe, expect, it } from 'vitest';

import { chartLabelPositions } from './label-positions';
import { newChartSpace } from './edit/new-chart';
import { setChartDataLabels } from './edit/legend-labels';
import { setChartPointLabel } from './edit/point-label';
import { setChartSeriesType } from './edit/series-type';
import type { ChartTypeTarget } from './edit/type';

const sheet = {
	categories: ['A', 'B'],
	series: [
		{ name: 'S', values: [1, 2] },
		{ name: 'T', values: [2, 3] },
		{ name: 'U', values: [3, 4] },
	],
};
const chart = (target: ChartTypeTarget) =>
	newChartSpace(sheet, target, { data: 'literal', profile: 'minimal' });

describe('chartLabelPositions', () => {
	it('lists what Office offers for each family', () => {
		expect(chartLabelPositions('barChart', 'clustered')).toStrictEqual([
			'ctr',
			'inEnd',
			'inBase',
			'outEnd',
		]);
		expect(chartLabelPositions('barChart', 'stacked')).not.toContain('outEnd');
		expect(chartLabelPositions('barChart', 'percentStacked')).not.toContain('outEnd');
		for (const element of ['lineChart', 'stockChart', 'scatterChart', 'bubbleChart'])
			expect(chartLabelPositions(element)).toStrictEqual(['ctr', 'l', 'r', 't', 'b']);
		for (const element of ['pieChart', 'pie3DChart', 'ofPieChart'])
			expect(chartLabelPositions(element)).toStrictEqual(['ctr', 'inEnd', 'outEnd', 'bestFit']);
		for (const element of [
			'bar3DChart',
			'line3DChart',
			'areaChart',
			'area3DChart',
			'radarChart',
			'doughnutChart',
			'surfaceChart',
		])
			expect(chartLabelPositions(element)).toStrictEqual([]);
	});
});

describe('label edits refuse a position the group does not accept', () => {
	it('checks chart-wide labels against every group', () => {
		expect(() => setChartDataLabels(chart({ kind: 'line' }), { position: 'outEnd' })).toThrow(
			RangeError,
		);
		expect(
			setChartDataLabels(chart({ kind: 'line' }), { position: 't' }).plotArea.groups[0]!.dataLabels
				?.position,
		).toBe('t');
		const combo = setChartSeriesType(chart({ kind: 'bar' }), 1, { kind: 'line' });
		expect(() => setChartDataLabels(combo, { position: 'inEnd' })).toThrow(RangeError);
		expect(() => setChartDataLabels(combo, { position: 'ctr' })).not.toThrow();
	});

	it("checks a series' and a point's labels against the series' group", () => {
		const stacked = chart({ kind: 'bar', grouping: 'stacked' });
		expect(() => setChartDataLabels(stacked, { position: 'outEnd' }, { series: 1 })).toThrow(
			RangeError,
		);
		expect(() => setChartPointLabel(stacked, 0, 1, { position: 'outEnd' })).toThrow(RangeError);
		expect(() => setChartPointLabel(chart({ kind: 'area' }), 0, 1, { position: 'ctr' })).toThrow(
			RangeError,
		);
		expect(
			setChartPointLabel(stacked, 0, 1, { position: 'inBase' }).plotArea.groups[0]!.series[0]!
				.dataLabels?.labels[0]?.position,
		).toBe('inBase');
	});

	it('still lets labels return to the default position', () => {
		const space = chart({ kind: 'area' });
		expect(() => setChartDataLabels(space, { showValue: true, position: null })).not.toThrow();
	});
});
