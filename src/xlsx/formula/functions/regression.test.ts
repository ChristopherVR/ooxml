// Fits with closed-form answers (derived by hand and checked in the comments), plus the
// documented LINEST example {1,9,5,7} against {0,4,2,3}, which is y = 2x + 1.
import { describe, expect, it } from 'vitest';
import { calc, calcArray, E } from '../test-helpers.js';

const cells = (name: string, values: number[]) =>
	Object.fromEntries(values.map((v, i) => [`${name}${i + 1}`, v]));
const near = (actual: unknown, expected: number, digits = 8) =>
	expect(actual as number).toBeCloseTo(expected, digits);

describe('LINEST', () => {
	const doc = { ...cells('A', [1, 9, 5, 7]), ...cells('B', [0, 4, 2, 3]) };
	it('returns the slope then the intercept', () => {
		const [line] = calcArray('LINEST(A1:A4,B1:B4)', doc, 1, 2);
		near(line?.[0], 2);
		near(line?.[1], 1);
	});

	it('without x fits against 1, 2, 3...', () => {
		const [line] = calcArray('LINEST(A1:A3)', cells('A', [3, 5, 7]), 1, 2);
		near(line?.[0], 2);
		near(line?.[1], 1);
	});

	it('reports the five statistics rows', () => {
		// y = 2,4,5,4,5 on x = 1..5: m 0.6, b 2.2, SSreg 3.6, SSres 2.4, df 3.
		const data = { ...cells('A', [2, 4, 5, 4, 5]), ...cells('B', [1, 2, 3, 4, 5]) };
		const t = calcArray('LINEST(A1:A5,B1:B5,TRUE,TRUE)', data, 5, 2);
		near(t[0]?.[0], 0.6);
		near(t[0]?.[1], 2.2);
		near(t[1]?.[0], Math.sqrt(0.08), 8);
		near(t[1]?.[1], Math.sqrt(0.88), 8);
		near(t[2]?.[0], 0.6);
		near(t[2]?.[1], Math.sqrt(0.8), 8);
		near(t[3]?.[0], 4.5);
		near(t[3]?.[1], 3);
		near(t[4]?.[0], 3.6);
		near(t[4]?.[1], 2.4);
	});

	it('solves several variables, last variable first', () => {
		// y = 1 + 2*x1 + 3*x2
		const x1 = [1, 2, 3, 4, 5];
		const x2 = [2, 1, 4, 3, 6];
		const data = {
			...cells(
				'A',
				x1.map((v, i) => 1 + 2 * v + 3 * (x2[i] ?? 0)),
			),
			...cells('B', x1),
			...cells('C', x2),
		};
		const [line] = calcArray('LINEST(A1:A5,B1:C5)', data, 1, 3);
		near(line?.[0], 3, 6);
		near(line?.[1], 2, 6);
		near(line?.[2], 1, 6);
	});

	it('forces the line through the origin when const is FALSE', () => {
		const data = { ...cells('A', [2, 4, 6]), ...cells('B', [1, 2, 3]) };
		const [line] = calcArray('LINEST(A1:A3,B1:B3,FALSE)', data, 1, 2);
		near(line?.[0], 2);
		expect(line?.[1]).toBe(0);
	});

	it('is #REF! for a two-dimensional y and #NUM! for collinear columns', () => {
		expect(calc('LINEST(A1:B2)', { A1: 1, A2: 2, B1: 3, B2: 4 })).toEqual(E.REF);
		const dup = {
			...cells('A', [1, 2, 3, 5]),
			...cells('B', [1, 2, 3, 4]),
			...cells('C', [2, 4, 6, 8]),
		};
		expect(calc('INDEX(LINEST(A1:A4,B1:C4),1,1)', dup)).toEqual(E.NUM);
	});
});

describe('TREND', () => {
	const doc = { ...cells('A', [1, 9, 5, 7]), ...cells('B', [0, 4, 2, 3]) };
	it('predicts at new x values, keeping their shape', () => {
		near(calc('TREND(A1:A4,B1:B4,5)', doc), 11);
		const out = calcArray('TREND(A1:A4,B1:B4,D1:D2)', { ...doc, D1: 5, D2: 6 }, 2, 1);
		near(out[0]?.[0], 11);
		near(out[1]?.[0], 13);
	});

	it('returns the fitted values without new_x', () => {
		const out = calcArray('TREND(A1:A4,B1:B4)', doc, 4, 1);
		expect(out.map((r) => Math.round(r[0] as number))).toEqual([1, 9, 5, 7]);
	});

	it('honours const FALSE', () => {
		const data = { ...cells('A', [2, 4, 6]), ...cells('B', [1, 2, 3]) };
		near(calc('TREND(A1:A3,B1:B3,4,FALSE)', data), 8);
	});
});

describe('LOGEST and GROWTH', () => {
	// y = 2 * 3^x
	const data = { ...cells('A', [2, 6, 18, 54]), ...cells('B', [0, 1, 2, 3]) };
	it('LOGEST returns the base then the multiplier', () => {
		const [line] = calcArray('LOGEST(A1:A4,B1:B4)', data, 1, 2);
		near(line?.[0], 3, 8);
		near(line?.[1], 2, 8);
	});

	it('GROWTH extrapolates the exponential', () => {
		near(calc('GROWTH(A1:A4,B1:B4,4)', data), 162, 6);
	});

	it('rejects non-positive y values', () => {
		expect(
			calc('LOGEST(A1:A3,B1:B3)', { ...cells('A', [1, 0, 2]), ...cells('B', [1, 2, 3]) }),
		).toEqual(E.NUM);
	});
});
