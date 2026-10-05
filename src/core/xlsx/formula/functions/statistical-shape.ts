import { ERR, fail, Matrix } from '../values.js';
import { collectNumbers, num, numeric, spec } from './helpers.js';
import * as S from './stats-core.js';
import { gamma, gammaLn } from './stats-core.js';
import type { FunctionSpec } from './types.js';

const C = 'Statistical';

function moments(values: number[]): { n: number; m: number; s: number } {
	const n = values.length;
	return { n, m: S.mean(values), s: S.stdev(values, true) };
}

function frequency(data: number[], bins: number[]): Matrix {
	const order = bins.map((b, i) => ({ b, i })).sort((x, y) => x.b - y.b);
	const counts = new Array<number>(bins.length + 1).fill(0);
	for (const v of data) {
		const slot = order.find((o) => v <= o.b);
		counts[slot ? slot.i : bins.length] = (counts[slot ? slot.i : bins.length] ?? 0) + 1;
	}
	return new Matrix(counts.map((c) => [c]));
}

/** Means, moments, transformations and counting distributions. */
export const STATISTICAL_SHAPE: FunctionSpec[] = [
	spec(
		'GEOMEAN',
		C,
		'GEOMEAN(number1, ...)',
		'The geometric mean.',
		1,
		255,
		(args, ctx) => {
			const values = collectNumbers(ctx, args);
			if (values.length === 0 || values.some((v) => v <= 0)) fail(ERR.NUM);
			return Math.exp(values.reduce((a, v) => a + Math.log(v), 0) / values.length);
		},
		['any'],
	),
	spec(
		'HARMEAN',
		C,
		'HARMEAN(number1, ...)',
		'The harmonic mean.',
		1,
		255,
		(args, ctx) => {
			const values = collectNumbers(ctx, args);
			if (values.length === 0 || values.some((v) => v <= 0)) fail(ERR.NUM);
			return values.length / values.reduce((a, v) => a + 1 / v, 0);
		},
		['any'],
	),
	spec(
		'AVEDEV',
		C,
		'AVEDEV(number1, ...)',
		'The mean absolute deviation from the mean.',
		1,
		255,
		(args, ctx) => {
			const values = collectNumbers(ctx, args);
			const m = S.mean(values);
			if (values.length === 0) fail(ERR.NUM);
			return values.reduce((a, v) => a + Math.abs(v - m), 0) / values.length;
		},
		['any'],
	),
	spec(
		'DEVSQ',
		C,
		'DEVSQ(number1, ...)',
		'The sum of squared deviations from the mean.',
		1,
		255,
		(args, ctx) => {
			const values = collectNumbers(ctx, args);
			return values.length ? S.devsq(values) : fail(ERR.NUM);
		},
		['any'],
	),
	spec(
		'KURT',
		C,
		'KURT(number1, ...)',
		'The kurtosis of a data set.',
		1,
		255,
		(args, ctx) => {
			const values = collectNumbers(ctx, args);
			const { n, m, s } = moments(values);
			if (n < 4 || s === 0) fail(ERR.DIV0);
			const sum4 = values.reduce((a, v) => a + ((v - m) / s) ** 4, 0);
			return (
				((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * sum4 -
				(3 * (n - 1) ** 2) / ((n - 2) * (n - 3))
			);
		},
		['any'],
	),
	spec(
		'SKEW',
		C,
		'SKEW(number1, ...)',
		'The skewness of a distribution.',
		1,
		255,
		(args, ctx) => {
			const values = collectNumbers(ctx, args);
			const { n, m, s } = moments(values);
			if (n < 3 || s === 0) fail(ERR.DIV0);
			return (n / ((n - 1) * (n - 2))) * values.reduce((a, v) => a + ((v - m) / s) ** 3, 0);
		},
		['any'],
	),
	spec(
		'SKEW.P',
		C,
		'SKEW.P(number1, ...)',
		'The population skewness.',
		1,
		255,
		(args, ctx) => {
			const values = collectNumbers(ctx, args);
			const m = S.mean(values);
			const sd = S.stdev(values, false);
			if (sd === 0) fail(ERR.DIV0);
			return values.reduce((a, v) => a + ((v - m) / sd) ** 3, 0) / values.length;
		},
		['any'],
	),
	numeric(
		'STANDARDIZE',
		C,
		'STANDARDIZE(x, mean, standard_dev)',
		'A normalized value.',
		3,
		3,
		(x, m, s) => {
			if ((s ?? 0) <= 0) fail(ERR.NUM);
			return (x - (m ?? 0)) / (s ?? 1);
		},
	),
	spec(
		'TRIMMEAN',
		C,
		'TRIMMEAN(array, percent)',
		'The mean of the interior of a data set.',
		2,
		2,
		(args, ctx) => {
			const values = S.sorted(collectNumbers(ctx, [args[0] ?? null]));
			const percent = num(args[1]);
			if (percent < 0 || percent >= 1 || values.length === 0) fail(ERR.NUM);
			const k = Math.floor((values.length * percent) / 2);
			return S.mean(values.slice(k, values.length - k));
		},
		['any', 'value'],
	),
	spec(
		'FREQUENCY',
		C,
		'FREQUENCY(data_array, bins_array)',
		'How often values occur within ranges, as a vertical array.',
		2,
		2,
		(args, ctx) =>
			frequency(collectNumbers(ctx, [args[0] ?? null]), collectNumbers(ctx, [args[1] ?? null])),
		['any'],
	),
	numeric(
		'PERMUT',
		C,
		'PERMUT(number, number_chosen)',
		'The number of permutations.',
		2,
		2,
		(n, k) => {
			const nn = Math.trunc(n);
			const kk = Math.trunc(k ?? 0);
			if (nn < 0 || kk < 0 || kk > nn) fail(ERR.NUM);
			let out = 1;
			for (let i = 0; i < kk; i++) out *= nn - i;
			return out;
		},
	),
	numeric(
		'PERMUTATIONA',
		C,
		'PERMUTATIONA(number, number_chosen)',
		'Permutations with repetitions.',
		2,
		2,
		(n, k) => {
			if (n < 0 || (k ?? 0) < 0) fail(ERR.NUM);
			return Math.pow(Math.trunc(n), Math.trunc(k ?? 0));
		},
	),
	numeric('FISHER', C, 'FISHER(x)', 'The Fisher transformation.', 1, 1, (x) => {
		if (x <= -1 || x >= 1) fail(ERR.NUM);
		return 0.5 * Math.log((1 + x) / (1 - x));
	}),
	numeric('FISHERINV', C, 'FISHERINV(y)', 'The inverse of the Fisher transformation.', 1, 1, (y) =>
		Math.tanh(y),
	),
	numeric('GAMMA', C, 'GAMMA(number)', 'The gamma function value.', 1, 1, (x) => {
		if (x <= 0 && Number.isInteger(x)) fail(ERR.NUM);
		return gamma(x);
	}),
	numeric(
		'GAMMALN',
		C,
		'GAMMALN(x)',
		'The natural logarithm of the gamma function.',
		1,
		1,
		gammaLn,
	),
	numeric(
		'GAMMALN.PRECISE',
		C,
		'GAMMALN.PRECISE(x)',
		'The natural logarithm of the gamma function.',
		1,
		1,
		gammaLn,
	),
];
