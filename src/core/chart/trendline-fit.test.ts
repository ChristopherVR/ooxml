import { describe, expect, it } from 'vitest';
import { chartErrorBarAmount, chartSampleVariance } from './error-bar-amount';
import { fitChartTrendline } from './trendline-fit';

const xs = [1, 2, 3, 4];

describe('fitChartTrendline', () => {
	it('fits a linear trend and samples it over the extended range', () => {
		const fit = fitChartTrendline({ kind: 'linear', forward: 1, backward: 0.5 }, xs, [3, 5, 7, 9]);
		expect(fit.coefficients[0]).toBeCloseTo(1);
		expect(fit.coefficients[1]).toBeCloseTo(2);
		expect(fit.rSquared).toBeCloseTo(1);
		expect(fit.equation).toBe('y = 2x + 1');
		expect(fit.points[0]?.x).toBe(0.5);
		expect(fit.points.at(-1)?.x).toBe(5);
		// (5 - 0.5) * 4 = 18 steps, raised to the minimum of 20.
		expect(fit.points).toHaveLength(21);
		for (const point of fit.points) expect(point.y).toBeCloseTo(2 * point.x + 1);
	});

	it('forces a linear trend through a set intercept', () => {
		const fit = fitChartTrendline({ kind: 'linear', intercept: 0 }, xs, [2, 4, 6, 8]);
		expect(fit.coefficients).toEqual([0, 2]);
	});

	it('fits exponential, logarithmic, power and polynomial curves', () => {
		const exp = fitChartTrendline({ kind: 'exponential' }, xs, [2, 4, 8, 16]);
		expect(exp.coefficients[0]).toBeCloseTo(1);
		expect(exp.coefficients[1]).toBeCloseTo(Math.LN2);
		const log = fitChartTrendline(
			{ kind: 'logarithmic' },
			xs,
			xs.map((x) => 3 * Math.log(x) + 1),
		);
		expect(log.coefficients[0]).toBeCloseTo(1);
		expect(log.coefficients[1]).toBeCloseTo(3);
		const power = fitChartTrendline(
			{ kind: 'power' },
			xs,
			xs.map((x) => 2 * x ** 3),
		);
		expect(power.coefficients[0]).toBeCloseTo(2);
		expect(power.coefficients[1]).toBeCloseTo(3);
		const poly = fitChartTrendline({ kind: 'polynomial', order: 2 }, xs, [1, 4, 9, 16]);
		expect(poly.coefficients.map((c) => Math.round(c * 1e6) / 1e6)).toEqual([0, 0, 1]);
		expect(poly.rSquared).toBeCloseTo(1);
	});

	it('averages full periods at the data x values', () => {
		const fit = fitChartTrendline({ kind: 'movingAvg', period: 3 }, xs, [3, 6, 9, 12]);
		expect(fit.points).toEqual([
			{ x: 3, y: 6 },
			{ x: 4, y: 9 },
		]);
		expect(fit.equation).toBe('3-period moving average');
	});

	it('returns no points for sparse data or a curve the data cannot fit', () => {
		expect(fitChartTrendline({ kind: 'linear' }, [1], [1]).points).toEqual([]);
		expect(fitChartTrendline({ kind: 'exponential' }, xs, [-1, -2, -3, 4]).points).toEqual([]);
	});
});

describe('fitChartTrendline over many points', () => {
	it('samples its range from more x values than a call can pass as arguments', () => {
		// A shuffled 1 to 2.99999, so the extremes sit inside the list rather than at its ends.
		const x = Array.from({ length: 200_000 }, (_, i) => 1 + ((i * 7919) % 200_000) / 100_000);
		const fit = fitChartTrendline(
			{ kind: 'linear' },
			x,
			x.map((v) => 2 * v + 1),
		);
		expect(fit.coefficients[1]).toBeCloseTo(2);
		expect(fit.points[0]?.x).toBe(1);
		expect(fit.points.at(-1)?.x).toBeCloseTo(2.99999);
	});
});

describe('chartErrorBarAmount', () => {
	const values = [2, 4, 4, 4, 5, 5, 7, 9];

	it('uses the sample variance', () => {
		expect(chartSampleVariance(values)).toBeCloseTo(32 / 7);
		expect(chartSampleVariance([5])).toBe(0);
		expect(chartSampleVariance([])).toBe(0);
	});

	it('computes each value type', () => {
		const amount = (options: Parameters<typeof chartErrorBarAmount>[0], side: 'plus' | 'minus') =>
			chartErrorBarAmount(options, values, 1, 2, side);
		expect(amount({ valueType: 'fixedVal', value: 1.5 }, 'plus')).toBe(1.5);
		expect(amount({ valueType: 'percentage', value: 25 }, 'plus')).toBe(1);
		expect(amount({ valueType: 'stdDev', value: 2 }, 'plus')).toBeCloseTo(2 * Math.sqrt(32 / 7));
		expect(amount({ valueType: 'stdErr' }, 'plus')).toBeCloseTo(Math.sqrt(32 / 7 / 8));
		const custom = { valueType: 'cust', customPlus: [0, 0, 3], customMinus: [0, 0, 1] } as const;
		expect(amount(custom, 'plus')).toBe(3);
		expect(amount(custom, 'minus')).toBe(1);
		expect(amount({ valueType: 'cust' }, 'minus')).toBe(0);
	});
});
