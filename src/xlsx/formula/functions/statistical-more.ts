import type { CallContext } from '../context.js';
import { ERR, fail, isError, Matrix, type Value } from '../values.js';
import { collectNumbers, num, spec } from './helpers.js';
import * as S from './stats-core.js';
import type { FunctionSpec } from './types.js';
import { STATISTICAL_SHAPE } from './statistical-shape.js';

const C = 'Statistical';

/** Pairs of numbers at the same positions of two arrays (pairs with a non-number are skipped). */
export function pairedNumbers(
	ctx: CallContext,
	a: Value,
	b: Value,
): { xs: number[]; ys: number[] } {
	const ma = ctx.toMatrix(a).flat();
	const mb = ctx.toMatrix(b).flat();
	if (ma.length !== mb.length) fail(ERR.NA);
	const xs: number[] = [];
	const ys: number[] = [];
	for (let i = 0; i < ma.length; i++) {
		const x = ma[i] ?? null;
		const y = mb[i] ?? null;
		if (isError(x)) fail(x);
		if (isError(y)) fail(y);
		if (typeof x === 'number' && typeof y === 'number') {
			xs.push(x);
			ys.push(y);
		}
	}
	return { xs, ys };
}

const pairFn = (
	name: string,
	description: string,
	syntax: string,
	fn: (xs: number[], ys: number[]) => number,
): FunctionSpec =>
	spec(
		name,
		C,
		syntax,
		description,
		2,
		2,
		(args, ctx) => {
			const { xs, ys } = pairedNumbers(ctx, args[0] ?? null, args[1] ?? null);
			return fn(xs, ys);
		},
		['any'],
	);

export const STATISTICAL_MORE: FunctionSpec[] = [
	...STATISTICAL_SHAPE,
	spec(
		'MODE.MULT',
		C,
		'MODE.MULT(number1, [number2], ...)',
		'A vertical array of the most frequent values.',
		1,
		255,
		(args, ctx) => new Matrix(S.modes(collectNumbers(ctx, args)).map((v) => [v])),
		['any'],
	),
	pairFn(
		'CORREL',
		'The correlation coefficient of two data sets.',
		'CORREL(array1, array2)',
		S.correlation,
	),
	pairFn(
		'PEARSON',
		'The Pearson correlation coefficient.',
		'PEARSON(array1, array2)',
		S.correlation,
	),
	pairFn(
		'RSQ',
		'The square of the Pearson correlation coefficient.',
		'RSQ(known_ys, known_xs)',
		(ys, xs) => S.correlation(xs, ys) ** 2,
	),
	pairFn('COVAR', 'The population covariance.', 'COVAR(array1, array2)', (x, y) =>
		S.covariance(x, y, false),
	),
	pairFn('COVARIANCE.P', 'The population covariance.', 'COVARIANCE.P(array1, array2)', (x, y) =>
		S.covariance(x, y, false),
	),
	pairFn('COVARIANCE.S', 'The sample covariance.', 'COVARIANCE.S(array1, array2)', (x, y) =>
		S.covariance(x, y, true),
	),
	pairFn(
		'SLOPE',
		'The slope of the linear regression line.',
		'SLOPE(known_ys, known_xs)',
		(ys, xs) => S.linearFit(xs, ys).slope,
	),
	pairFn(
		'INTERCEPT',
		'The intercept of the linear regression line.',
		'INTERCEPT(known_ys, known_xs)',
		(ys, xs) => S.linearFit(xs, ys).intercept,
	),
	pairFn(
		'STEYX',
		'The standard error of the predicted y values.',
		'STEYX(known_ys, known_xs)',
		(ys, xs) => {
			const n = xs.length;
			if (n < 3) fail(ERR.DIV0);
			const mx = S.mean(xs);
			const my = S.mean(ys);
			let sxy = 0;
			let sxx = 0;
			let syy = 0;
			for (let i = 0; i < n; i++) {
				const dx = (xs[i] as number) - mx;
				const dy = (ys[i] as number) - my;
				sxy += dx * dy;
				sxx += dx * dx;
				syy += dy * dy;
			}
			if (sxx === 0) fail(ERR.DIV0);
			return Math.sqrt((syy - (sxy * sxy) / sxx) / (n - 2));
		},
	),
	...(['FORECAST', 'FORECAST.LINEAR'] as const).map((name) =>
		spec(
			name,
			C,
			`${name}(x, known_ys, known_xs)`,
			'Predicts a value along a linear trend.',
			3,
			3,
			(args, ctx) => {
				const x = num(args[0]);
				const { xs, ys } = pairedNumbers(ctx, args[2] ?? null, args[1] ?? null);
				const fit = S.linearFit(xs, ys);
				return fit.intercept + fit.slope * x;
			},
			['value', 'any'],
		),
	),
];
