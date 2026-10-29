// FORECAST.ETS family: structural behaviour, argument validation and series whose answer is
// determined (an exact linear trend, an exactly repeating pattern). Smoothed series are compared
// with Excel's recorded answers in ../forecast-ets-excel.test.ts.
import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers';

const col = (name: string, values: number[]) =>
	Object.fromEntries(values.map((v, i) => [`${name}${i + 1}`, v]));
const near = (actual: unknown, expected: number, digits = 6) =>
	expect(actual as number).toBeCloseTo(expected, digits);

const linear = { ...col('A', [3, 5, 7, 9, 11, 13]), ...col('B', [1, 2, 3, 4, 5, 6]) };
const pattern = [10, 20, 15, 30];
const seasonal = {
	...col(
		'A',
		Array.from({ length: 12 }, (_v, i) => pattern[i % 4] as number),
	),
	...col(
		'B',
		Array.from({ length: 12 }, (_v, i) => i + 1),
	),
};

describe('FORECAST.ETS', () => {
	it('extends an exact linear trend', () => {
		near(calc('FORECAST.ETS(7,A1:A6,B1:B6,0)', linear), 15, 4);
		near(calc('FORECAST.ETS(10,A1:A6,B1:B6,0)', linear), 21, 4);
	});

	it('repeats an exactly repeating pattern', () => {
		near(calc('FORECAST.ETS(13,A1:A12,B1:B12,4)', seasonal), 10, 4);
		near(calc('FORECAST.ETS(16,A1:A12,B1:B12,4)', seasonal), 30, 4);
	});

	it('reads fitted values inside the history', () => {
		near(calc('FORECAST.ETS(3,A1:A6,B1:B6,0)', linear), 7, 4);
	});

	it('does not depend on the order of the timeline', () => {
		const shuffled = { ...col('A', [13, 3, 9, 5, 11, 7]), ...col('B', [6, 1, 4, 2, 5, 3]) };
		near(calc('FORECAST.ETS(7,A1:A6,B1:B6,0)', shuffled), 15, 4);
	});

	it('completes a missing timeline point by interpolation', () => {
		// Points at 1, 2, 4, 5 on a linear series; 3 is interpolated to 7.
		const gap = { ...col('A', [3, 5, 9, 11]), ...col('B', [1, 2, 4, 5]) };
		near(calc('FORECAST.ETS(6,A1:A4,B1:B4,0,1)', gap), 13, 4);
	});

	it('aggregates duplicate timestamps (average by default, sum on request)', () => {
		const dup = { ...col('A', [3, 4, 6, 7, 9, 11]), ...col('B', [1, 2, 2, 3, 4, 5]) };
		near(calc('FORECAST.ETS(6,A1:A6,B1:B6,0)', dup), 13, 4);
		// Summed per timestamp the series is 2, 4, 6, 8, 10, an exact line that continues to 12.
		const sum = { ...col('A', [1, 1, 4, 6, 8, 10]), ...col('B', [1, 1, 2, 3, 4, 5]) };
		near(calc('FORECAST.ETS(6,A1:A6,B1:B6,0,1,7)', sum), 12, 4);
		// Counted per timestamp it is 2, 1, 1, 1, 1, so the forecast is far from 12.
		expect(calc('FORECAST.ETS(6,A1:A6,B1:B6,0,1,2)', sum) as number).toBeLessThan(3);
	});

	it('validates its arguments', () => {
		expect(calc('FORECAST.ETS(7,A1:A6,B1:B5,0)', linear)).toEqual(E.NA);
		expect(calc('FORECAST.ETS(7,A1:A6,B1:B6,-1)', linear)).toEqual(E.NUM);
		expect(calc('FORECAST.ETS(7,A1:A6,B1:B6,1.5)', linear)).toEqual(E.NUM);
		expect(calc('FORECAST.ETS(7,A1:A6,B1:B6,9000)', linear)).toEqual(E.NUM);
		expect(calc('FORECAST.ETS(7,A1:A6,B1:B6,0,2)', linear)).toEqual(E.NUM);
		expect(calc('FORECAST.ETS(7,A1:A6,B1:B6,0,1,8)', linear)).toEqual(E.NUM);
		expect(calc('FORECAST.ETS(0,A1:A6,B1:B6,0)', linear)).toEqual(E.NUM);
		expect(calc('FORECAST.ETS(7,A1:A6,B1:B6,0)', { ...linear, A2: 'x' })).toEqual(E.VALUE);
	});

	it('rejects a timeline without a constant step', () => {
		const uneven = { ...col('A', [1, 2, 3, 4]), ...col('B', [1, 2, 4, 5.5]) };
		expect(calc('FORECAST.ETS(7,A1:A4,B1:B4,0)', uneven)).toEqual(E.NUM);
	});
});

describe('FORECAST.ETS.SEASONALITY', () => {
	it('finds the length of a repeating pattern', () => {
		const wave = [0, 10, 0, -10];
		const cells = {
			...col(
				'A',
				Array.from({ length: 12 }, (_v, i) => wave[i % 4] as number),
			),
			...col(
				'B',
				Array.from({ length: 12 }, (_v, i) => i + 1),
			),
		};
		expect(calc('FORECAST.ETS.SEASONALITY(A1:A12,B1:B12)', cells)).toBe(4);
	});

	it('reports the dominant harmonic like Excel (a 10,20,15,30 pattern reports 2)', () => {
		expect(calc('FORECAST.ETS.SEASONALITY(A1:A12,B1:B12)', seasonal)).toBe(2);
	});

	it('is 0 for a straight line', () => {
		expect(calc('FORECAST.ETS.SEASONALITY(A1:A6,B1:B6)', linear)).toBe(0);
	});
});

describe('FORECAST.ETS.CONFINT', () => {
	it('is zero for a perfectly fitted series and non-negative otherwise', () => {
		near(calc('FORECAST.ETS.CONFINT(8,A1:A6,B1:B6,0.95,0)', linear), 0, 4);
		const noisy = { ...col('A', [3, 6, 6, 10, 10, 14]), ...col('B', [1, 2, 3, 4, 5, 6]) };
		const near7 = calc('FORECAST.ETS.CONFINT(7,A1:A6,B1:B6,0.95,0)', noisy) as number;
		const far = calc('FORECAST.ETS.CONFINT(12,A1:A6,B1:B6,0.95,0)', noisy) as number;
		expect(near7).toBeGreaterThan(0);
		expect(far).toBeGreaterThanOrEqual(near7);
	});

	it('widens with the confidence level and validates it', () => {
		const noisy = { ...col('A', [3, 6, 6, 10, 10, 14]), ...col('B', [1, 2, 3, 4, 5, 6]) };
		const p90 = calc('FORECAST.ETS.CONFINT(8,A1:A6,B1:B6,0.9,0)', noisy) as number;
		const p99 = calc('FORECAST.ETS.CONFINT(8,A1:A6,B1:B6,0.99,0)', noisy) as number;
		expect(p99).toBeGreaterThan(p90);
		expect(calc('FORECAST.ETS.CONFINT(8,A1:A6,B1:B6,1,0)', linear)).toEqual(E.NUM);
		expect(calc('FORECAST.ETS.CONFINT(8,A1:A6,B1:B6,0,0)', linear)).toEqual(E.NUM);
	});
});

describe('FORECAST.ETS.STAT', () => {
	it('reports the step size and zero error for an exact fit', () => {
		expect(calc('FORECAST.ETS.STAT(A1:A6,B1:B6,8,0)', linear)).toBe(1);
		near(calc('FORECAST.ETS.STAT(A1:A6,B1:B6,6,0)', linear), 0, 6);
		near(calc('FORECAST.ETS.STAT(A1:A6,B1:B6,7,0)', linear), 0, 6);
		near(calc('FORECAST.ETS.STAT(A1:A6,B1:B6,5,0)', linear), 0, 6);
	});

	it('keeps the smoothing parameters within 0 and 1', () => {
		for (const kind of [1, 2, 3]) {
			const v = calc(`FORECAST.ETS.STAT(A1:A12,B1:B12,${kind})`, seasonal) as number;
			expect(v).toBeGreaterThanOrEqual(0);
			expect(v).toBeLessThanOrEqual(1);
		}
	});

	it('rejects an unknown statistic type', () => {
		expect(calc('FORECAST.ETS.STAT(A1:A6,B1:B6,0)', linear)).toEqual(E.NUM);
		expect(calc('FORECAST.ETS.STAT(A1:A6,B1:B6,9)', linear)).toEqual(E.NUM);
	});
});
