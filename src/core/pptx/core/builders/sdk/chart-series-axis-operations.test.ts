import { describe, expect, it } from 'vitest';

import type { PptxChartData } from '../../types';
import {
	chartSupportsSecondaryAxis,
	isChartSeriesOnSecondaryAxis,
	withChartSeriesOnSecondaryAxis,
} from './chart-series-axis-operations';

const chart = (overrides: Partial<PptxChartData> = {}): PptxChartData => ({
	chartType: 'bar',
	categories: ['Q1', 'Q2'],
	series: [
		{ name: 'East', values: [1, 2], axisId: 20 },
		{ name: 'West', values: [300, 400], axisId: 20 },
	],
	axes: [
		{ axisType: 'catAx', axisId: 10, crossAxisId: 20, axPos: 'b' },
		{ axisType: 'valAx', axisId: 20, crossAxisId: 10, axPos: 'l' },
	],
	...overrides,
});

describe('withChartSeriesOnSecondaryAxis', () => {
	it('creates a secondary value axis and its hidden category axis', () => {
		const next = withChartSeriesOnSecondaryAxis(chart(), 1, true);
		expect(next.axes).toHaveLength(4);
		expect(next.axes?.[2]).toMatchObject({
			axisType: 'valAx',
			axisId: 21,
			crossAxisId: 22,
			axPos: 'r',
			crosses: 'max',
		});
		expect(next.axes?.[3]).toMatchObject({
			axisType: 'catAx',
			axisId: 22,
			crossAxisId: 21,
			deleted: true,
		});
		expect(next.series[1].axisId).toBe(21);
		expect(isChartSeriesOnSecondaryAxis(next, 1)).toBe(true);
		expect(isChartSeriesOnSecondaryAxis(next, 0)).toBe(false);
	});

	it('reuses an existing secondary axis and drops it with the last series', () => {
		const one = withChartSeriesOnSecondaryAxis(chart(), 1, true);
		const two = withChartSeriesOnSecondaryAxis(one, 0, true);
		expect(two.axes).toHaveLength(4);
		const back = withChartSeriesOnSecondaryAxis(
			withChartSeriesOnSecondaryAxis(two, 0, false),
			1,
			false,
		);
		expect(back.axes).toHaveLength(2);
		expect(back.series.map((s) => s.axisId)).toEqual([20, 20]);
	});

	it('renumbers axis format keys after removing the secondary pair', () => {
		const one = withChartSeriesOnSecondaryAxis(chart(), 1, true);
		const formatted: PptxChartData = {
			...one,
			elementFormats: {
				'axis:2': { line: { kind: 'none' } },
				'axis:1': { line: { kind: 'none' } },
			},
		};
		const back = withChartSeriesOnSecondaryAxis(formatted, 1, false);
		expect(Object.keys(back.elementFormats ?? {})).toEqual(['axis:1']);
	});

	it('leaves charts without axes alone', () => {
		const pie = chart({ chartType: 'pie', axes: [] });
		expect(chartSupportsSecondaryAxis(pie)).toBe(false);
		expect(withChartSeriesOnSecondaryAxis(pie, 0, true)).toBe(pie);
	});
});
