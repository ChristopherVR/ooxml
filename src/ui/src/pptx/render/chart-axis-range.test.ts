import type { PptxChartSeries } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	computeStackedValueRangeForAxis,
	computeValueRangeForAxis,
	computeValueRangeForChart,
	generateAxisTicks,
	generateLogTicks,
	generateMinorAxisTicks,
} from './chart-axis';

const SERIES: PptxChartSeries[] = [{ name: 'Values', values: [20, 40] }];

describe('stacked value ranges on an axis', () => {
	const STACKED: PptxChartSeries[] = [
		{ name: 'A', values: [20, 10] },
		{ name: 'B', values: [30, 40] },
	];

	it('scales the category totals by explicit bounds', () => {
		expect(
			computeStackedValueRangeForAxis(STACKED, 2, { axisType: 'valAx', min: 0, max: 100 }, false),
		).toStrictEqual({ min: 0, max: 100, span: 100 });
	});

	it('keeps the automatic scale of the totals without bounds', () => {
		const range = computeStackedValueRangeForAxis(STACKED, 2, { axisType: 'valAx' }, false);
		expect(range.min).toBe(0);
		expect(range.max).toBeGreaterThanOrEqual(50);
		expect(range.majorUnit).toBeGreaterThan(0);
	});

	it('keeps percentStacked at 0 to 100 without negative values and follows c:orientation', () => {
		const axis = { axisType: 'valAx', min: 0, max: 0.5, orientation: 'maxMin' } as const;
		expect(computeStackedValueRangeForAxis(STACKED, 2, axis, true)).toStrictEqual({
			min: 0,
			max: 100,
			span: 100,
			reverseOrder: true,
		});
	});

	it('extends percentStacked below zero to fit negative shares', () => {
		// Category shares: 75% and -25%, then -100% for the all-negative one.
		const mixed: PptxChartSeries[] = [
			{ name: 'A', values: [3, -2] },
			{ name: 'B', values: [-1, -2] },
		];
		expect(computeStackedValueRangeForAxis(mixed.slice(0, 1), 1, undefined, true)).toMatchObject({
			min: 0,
			max: 100,
		});
		const partly = computeStackedValueRangeForAxis(mixed, 1, undefined, true);
		expect(partly.max).toBe(100);
		expect(partly.min).toBeLessThanOrEqual(-25);
		expect(partly.min).toBeGreaterThan(-100);
		expect(partly.span).toBe(100 - partly.min);
		expect(computeStackedValueRangeForAxis(mixed, 2, undefined, true)).toMatchObject({
			min: -100,
			max: 100,
			span: 200,
		});
	});
});

describe('axis-constrained value ranges', () => {
	it('uses explicit linear minimum and maximum bounds', () => {
		expect(computeValueRangeForAxis(SERIES, { axisType: 'valAx', min: 10, max: 50 })).toStrictEqual(
			{ min: 10, max: 50, span: 40 },
		);
	});

	it('uses explicit bounds with logarithmic span math', () => {
		const range = computeValueRangeForAxis([{ name: 'Log', values: [2, 200] }], {
			axisType: 'valAx',
			min: 1,
			max: 1000,
			logScale: true,
			logBase: 10,
		});
		expect(range).toMatchObject({ min: 1, max: 1000, logScale: true, logBase: 10 });
		expect(range.span).toBeCloseTo(3);
	});

	it('does not apply a secondary log axis to the primary range', () => {
		const range = computeValueRangeForChart(SERIES, [
			{ axisType: 'valAx', axPos: 'l', min: 0, max: 100 },
			{ axisType: 'valAx', axPos: 'r', logScale: true, logBase: 10 },
		]);
		expect(range).toStrictEqual({ min: 0, max: 100, span: 100 });
	});

	it('keeps logarithmic ticks inside explicit non-power bounds', () => {
		expect(
			generateLogTicks({ min: 2, max: 200, span: 2, logScale: true, logBase: 10 }),
		).toStrictEqual([10, 100]);
	});

	it('retains exact power ticks across negative and positive exponents', () => {
		expect(
			generateLogTicks({ min: 0.01, max: 100_000, span: 7, logScale: true, logBase: 10 }),
		).toStrictEqual([0.01, 0.1, 1, 10, 100, 1_000, 10_000, 100_000]);
	});

	it('retains reversed axis direction on the computed range', () => {
		expect(
			computeValueRangeForAxis(SERIES, { axisType: 'valAx', orientation: 'maxMin' }),
		).toMatchObject({ reverseOrder: true });
	});

	it('uses explicit major and minor units without duplicating major ticks', () => {
		const range = { min: 0, max: 100, span: 100 };
		const axis = { axisType: 'valAx' as const, majorUnit: 20, minorUnit: 10 };
		expect(generateAxisTicks(range, axis, 5)).toStrictEqual([0, 20, 40, 60, 80, 100]);
		expect(generateMinorAxisTicks(range, axis)).toStrictEqual([10, 30, 50, 70, 90]);
	});
});
