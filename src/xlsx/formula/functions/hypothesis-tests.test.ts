// Expected values are the worked examples in Microsoft's function reference for each function.
import { describe, expect, it } from 'vitest';
import { calc, E } from '../test-helpers.js';

const close = (formula: string, expected: number, digits: number, cells = {}) =>
	expect(calc(formula, cells) as number).toBeCloseTo(expected, digits);

describe('Z.TEST', () => {
	const cells = Object.fromEntries([3, 6, 7, 8, 6, 5, 4, 2, 1, 9].map((v, i) => [`A${i + 1}`, v]));
	it('uses the sample deviation unless sigma is given', () => {
		close('Z.TEST(A1:A10,4)', 0.090574, 6, cells);
		close('ZTEST(A1:A10,6)', 0.863043, 6, cells);
		close('Z.TEST(A1:A10,6,1)', 0.997787, 6, cells);
	});
	it('rejects an empty sample and a non-positive sigma', () => {
		expect(calc('Z.TEST(A1:A3,4)')).toEqual(E.NA);
		expect(calc('Z.TEST(A1:A10,4,0)', cells)).toEqual(E.NUM);
	});
});

describe('T.TEST', () => {
	const a = [3, 4, 5, 8, 9, 1, 2, 4, 5];
	const b = [6, 19, 3, 2, 14, 4, 5, 17, 1];
	const cells = {
		...Object.fromEntries(a.map((v, i) => [`A${i + 1}`, v])),
		...Object.fromEntries(b.map((v, i) => [`B${i + 1}`, v])),
	};
	it('paired, equal-variance and Welch tests', () => {
		close('T.TEST(A1:A9,B1:B9,2,1)', 0.196016, 6, cells);
		close('TTEST(A1:A9,B1:B9,2,2)', 0.2, 1, cells);
		const welch = calc('T.TEST(A1:A9,B1:B9,2,3)', cells) as number;
		expect(welch).toBeGreaterThan(0.1);
		expect(welch).toBeLessThan(0.3);
		const one = calc('T.TEST(A1:A9,B1:B9,1,1)', cells) as number;
		expect(one * 2).toBeCloseTo(0.196016, 6);
	});
	it('validates tails, type and paired sizes', () => {
		expect(calc('T.TEST(A1:A9,B1:B9,3,1)', cells)).toEqual(E.NUM);
		expect(calc('T.TEST(A1:A9,B1:B9,2,4)', cells)).toEqual(E.NUM);
		expect(calc('T.TEST(A1:A9,B1:B8,2,1)', cells)).toEqual(E.NA);
	});
});

describe('F.TEST', () => {
	const cells = {
		...Object.fromEntries([6, 7, 9, 15, 21].map((v, i) => [`A${i + 1}`, v])),
		...Object.fromEntries([20, 28, 31, 38, 40].map((v, i) => [`B${i + 1}`, v])),
	};
	it('is the two-tailed probability that the variances differ', () => {
		close('F.TEST(A1:A5,B1:B5)', 0.648318, 6, cells);
		close('FTEST(A1:A5,B1:B5)', 0.648318, 6, cells);
	});
	it('needs two points and a non-zero variance in each array', () => {
		expect(calc('F.TEST(A1,B1:B5)', cells)).toEqual(E.DIV0);
	});
});

describe('CHISQ.TEST', () => {
	const cells = {
		A1: 58,
		B1: 35,
		A2: 11,
		B2: 25,
		A3: 10,
		B3: 23,
		D1: 45.35,
		E1: 47.65,
		D2: 17.56,
		E2: 18.44,
		D3: 16.09,
		E3: 16.91,
	};
	it('matches the documented contingency example', () => {
		close('CHISQ.TEST(A1:B3,D1:E3)', 0.0003082, 7, cells);
		close('CHITEST(A1:B3,D1:E3)', 0.0003082, 7, cells);
	});
	it('rejects mismatched shapes and zero expectations', () => {
		expect(calc('CHISQ.TEST(A1:B3,D1:E2)', cells)).toEqual(E.NA);
		expect(calc('CHISQ.TEST(A1:B3,G1:H3)', cells)).toEqual(E.DIV0);
	});
});

describe('PROB', () => {
	const cells = { A1: 0, A2: 1, A3: 2, A4: 3, B1: 0.2, B2: 0.3, B3: 0.1, B4: 0.4 };
	it('sums the probabilities inside the limits', () => {
		close('PROB(A1:A4,B1:B4,2)', 0.1, 10, cells);
		close('PROB(A1:A4,B1:B4,1,3)', 0.8, 10, cells);
	});
	it('requires probabilities that sum to one', () => {
		expect(calc('PROB(A1:A4,B1:B3,1)', cells)).toEqual(E.NA);
		expect(calc('PROB(A1:A3,B1:B3,1)', cells)).toEqual(E.NUM);
	});
});
