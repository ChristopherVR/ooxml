import type { CallContext } from '../context.js';
import { ERR, fail, isError, Matrix, type Scalar, type Value } from '../values.js';
import { bool, omitted, spec } from './helpers.js';
import { fitLeastSquares, fitStatistics, type Fit } from './regression-core.js';
import type { FunctionSpec } from './types.js';

const C = 'Statistical';

/** `known_y` and `known_x` laid out as observations (rows) by variables (columns). */
interface Layout {
	y: number[];
	x: number[][];
	/** Whether the observations ran down a column (the result then runs down one too). */
	column: boolean;
}

function numbersOf(ctx: CallContext, value: Value | undefined): number[][] {
	return ctx.toMatrix(value ?? null).data.map((row) =>
		row.map((v) => {
			if (isError(v)) fail(v);
			if (typeof v !== 'number') fail(ERR.VALUE);
			return v;
		}),
	);
}

/** Observations as rows: a single-row range is turned on its side. */
const asObservations = (m: number[][], column: boolean): number[][] =>
	column ? m : (m[0] ?? []).map((_, c) => m.map((row) => row[c] ?? 0));

function layout(ctx: CallContext, args: Value[]): Layout {
	const ym = numbersOf(ctx, args[0]);
	const rows = ym.length;
	const cols = ym[0]?.length ?? 0;
	if (rows === 0 || cols === 0 || (rows > 1 && cols > 1)) fail(ERR.REF);
	const column = cols === 1;
	const y = ym.flat();
	if (omitted(args, 1)) return { y, x: y.map((_, i) => [i + 1]), column };
	const xm = numbersOf(ctx, args[1]);
	const x = asObservations(xm, column);
	if (x.length !== y.length) {
		// A single variable may be given in any shape of the same size.
		if (xm.flat().length !== y.length) fail(ERR.REF);
		return { y, x: xm.flat().map((v) => [v]), column };
	}
	return { y, x, column };
}

const withConst = (args: Value[], index: number): boolean =>
	omitted(args, index) ? true : bool(args[index]);

const row = (values: Scalar[], width: number): Scalar[] => [
	...values,
	...new Array<Scalar>(Math.max(0, width - values.length)).fill(ERR.NA),
];

/** LINEST / LOGEST output: coefficients in reverse variable order, then the constant. */
function coefficientTable(fit: Fit, stats: boolean, transform: (v: number) => number): Matrix {
	const k = fit.slopes.length;
	const width = k + 1;
	const slopes = [...fit.slopes].reverse().map(transform);
	const head = [...slopes, transform(fit.intercept)];
	if (!stats) return new Matrix([head]);
	const s = fitStatistics(fit);
	const errors = [...s.slopeErrors].reverse();
	return new Matrix([
		head,
		row([...errors, s.interceptError ?? ERR.NA], width),
		row([s.rSquared, s.standardError], width),
		row([s.fStatistic === Infinity ? ERR.NUM : s.fStatistic, s.df], width),
		row([s.ssRegression, s.ssResidual], width),
	]);
}

/** The predictions of a fit at new observations, shaped like Excel's TREND. */
function predictions(
	ctx: CallContext,
	args: Value[],
	lay: Layout,
	fit: Fit,
	transform: (v: number) => number,
): Matrix {
	const predict = (obs: number[]): number =>
		transform(fit.intercept + fit.slopes.reduce((a, m, j) => a + m * (obs[j] ?? 0), 0));
	if (omitted(args, 2)) {
		const out = lay.x.map(predict);
		return new Matrix(lay.column ? out.map((v) => [v]) : [out]);
	}
	const target = numbersOf(ctx, args[2]);
	if (fit.slopes.length === 1) return new Matrix(target.map((r) => r.map((v) => predict([v]))));
	const observations = asObservations(target, lay.column);
	if ((observations[0]?.length ?? 0) !== fit.slopes.length) fail(ERR.REF);
	const out = observations.map(predict);
	return new Matrix(lay.column ? out.map((v) => [v]) : [out]);
}

const logs = (lay: Layout): Layout => ({
	...lay,
	y: lay.y.map((v) => (v > 0 ? Math.log(v) : fail(ERR.NUM))),
});

const ARGS = ['any'] as const;

/** Multiple linear and exponential regression. */
export const REGRESSION_FUNCTIONS: FunctionSpec[] = [
	spec(
		'LINEST',
		C,
		'LINEST(known_y, [known_x], [const], [stats])',
		'The statistics of a straight-line (or multiple linear) least-squares fit.',
		1,
		4,
		(args, ctx) => {
			const lay = layout(ctx, args);
			const fit = fitLeastSquares(lay.x, lay.y, withConst(args, 2));
			return coefficientTable(fit, !omitted(args, 3) && bool(args[3]), (v) => v);
		},
		ARGS,
	),
	spec(
		'LOGEST',
		C,
		'LOGEST(known_y, [known_x], [const], [stats])',
		'The statistics of an exponential-curve fit.',
		1,
		4,
		(args, ctx) => {
			const lay = logs(layout(ctx, args));
			const fit = fitLeastSquares(lay.x, lay.y, withConst(args, 2));
			const stats = !omitted(args, 3) && bool(args[3]);
			const table = coefficientTable(fit, stats, Math.exp);
			if (!stats) return table;
			// Only the coefficients are back-transformed; the statistics describe the log fit.
			const plain = coefficientTable(fit, true, (v) => v).data;
			return new Matrix([table.data[0] ?? [], ...plain.slice(1)]);
		},
		ARGS,
	),
	spec(
		'TREND',
		C,
		'TREND(known_y, [known_x], [new_x], [const])',
		'Values along a linear trend.',
		1,
		4,
		(args, ctx) => {
			const lay = layout(ctx, args);
			const fit = fitLeastSquares(lay.x, lay.y, withConst(args, 3));
			return predictions(ctx, args, lay, fit, (v) => v);
		},
		ARGS,
	),
	spec(
		'GROWTH',
		C,
		'GROWTH(known_y, [known_x], [new_x], [const])',
		'Values along an exponential trend.',
		1,
		4,
		(args, ctx) => {
			const lay = logs(layout(ctx, args));
			const fit = fitLeastSquares(lay.x, lay.y, withConst(args, 3));
			return predictions(ctx, args, lay, fit, Math.exp);
		},
		ARGS,
	),
];
