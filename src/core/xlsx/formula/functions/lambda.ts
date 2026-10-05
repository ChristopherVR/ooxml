import type { CallContext } from '../context.js';
import { pick } from '../operators.js';
import { ERR, fail, LambdaValue, Matrix, RefValue, type Scalar, type Value } from '../values.js';
import { int, spec } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Logical';

function lambdaArg(value: Value | undefined): LambdaValue {
	if (!(value instanceof LambdaValue)) fail(ERR.VALUE);
	return value;
}

/** The scalar a lambda result contributes to an array (a 1x1 array or cell becomes its value). */
function element(ctx: CallContext, value: Value): Scalar {
	if (value instanceof Matrix) {
		if (value.rows !== 1 || value.cols !== 1) return ERR.CALC;
		return value.get(0, 0);
	}
	if (value instanceof RefValue) return value.isCell() ? ctx.toScalar(value) : ERR.CALC;
	if (value instanceof LambdaValue) return ERR.CALC;
	return value === null ? 0 : value;
}

export const LAMBDA_HELPERS: FunctionSpec[] = [
	spec(
		'MAP',
		C,
		'MAP(array1, ..., lambda)',
		'Applies a LAMBDA to each value of arrays.',
		2,
		254,
		(args, ctx) => {
			const fn = lambdaArg(args[args.length - 1]);
			const arrays = args.slice(0, -1).map((a) => ctx.toMatrix(a));
			const rows = arrays.reduce((n, m) => Math.max(n, m.rows), 0);
			const cols = arrays.reduce((n, m) => Math.max(n, m.cols), 0);
			return Matrix.build(rows, cols, (r, c) =>
				element(
					ctx,
					ctx.callLambda(
						fn,
						arrays.map((m) => pick(m, r, c)),
					),
				),
			);
		},
		['any'],
	),
	spec(
		'REDUCE',
		C,
		'REDUCE([initial_value], array, lambda)',
		'Reduces an array to an accumulated value.',
		2,
		3,
		(args, ctx) => {
			const fn = lambdaArg(args[args.length - 1]);
			const array = ctx.toMatrix(args.length === 3 ? (args[1] ?? null) : (args[0] ?? null));
			let acc: Value = args.length === 3 ? (args[0] ?? null) : null;
			for (const v of array.flat()) acc = ctx.callLambda(fn, [acc, v]);
			return acc;
		},
		['any'],
	),
	spec(
		'SCAN',
		C,
		'SCAN([initial_value], array, lambda)',
		'The intermediate values of a reduction, as an array.',
		2,
		3,
		(args, ctx) => {
			const fn = lambdaArg(args[args.length - 1]);
			const array = ctx.toMatrix(args.length === 3 ? (args[1] ?? null) : (args[0] ?? null));
			let acc: Value = args.length === 3 ? (args[0] ?? null) : null;
			return array.map((v) => {
				acc = ctx.callLambda(fn, [acc, v]);
				return element(ctx, acc);
			});
		},
		['any'],
	),
	spec(
		'BYROW',
		C,
		'BYROW(array, lambda)',
		'Applies a LAMBDA to each row, returning a column.',
		2,
		2,
		(args, ctx) => {
			const fn = lambdaArg(args[1]);
			const array = ctx.toMatrix(args[0] ?? null);
			return new Matrix(
				array.data.map((row) => [element(ctx, ctx.callLambda(fn, [new Matrix([row])]))]),
			);
		},
		['any'],
	),
	spec(
		'BYCOL',
		C,
		'BYCOL(array, lambda)',
		'Applies a LAMBDA to each column, returning a row.',
		2,
		2,
		(args, ctx) => {
			const fn = lambdaArg(args[1]);
			const array = ctx.toMatrix(args[0] ?? null);
			const out: Scalar[] = [];
			for (let c = 0; c < array.cols; c++) {
				const column = new Matrix(array.data.map((row) => [row[c] ?? null]));
				out.push(element(ctx, ctx.callLambda(fn, [column])));
			}
			return new Matrix([out]);
		},
		['any'],
	),
	spec(
		'MAKEARRAY',
		C,
		'MAKEARRAY(rows, columns, lambda)',
		'An array built by calling a LAMBDA for each position.',
		3,
		3,
		(args, ctx) => {
			const rows = int(ctx.toScalar(args[0] ?? null));
			const cols = int(ctx.toScalar(args[1] ?? null));
			if (rows < 1 || cols < 1) fail(ERR.VALUE);
			const fn = lambdaArg(args[2]);
			return Matrix.build(rows, cols, (r, c) => element(ctx, ctx.callLambda(fn, [r + 1, c + 1])));
		},
		['any'],
	),
];
