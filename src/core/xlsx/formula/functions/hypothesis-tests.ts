import type { CallContext } from '../context';
import { ERR, fail, isError, type Value } from '../values';
import { chiCdf, normCdf, tCdf } from './dist-core';
import { fCdf } from './distributions-f';
import { collectNumbers, int, num, spec } from './helpers';
import { pairedNumbers } from './statistical-more';
import * as S from './stats-core';
import type { FunctionSpec } from './types';

const C = 'Statistical';
const ARRAYS = ['any'] as const;

/** The numbers of one array argument (text and blanks skipped). */
const numbersOf = (ctx: CallContext, value: Value | undefined): number[] =>
	collectNumbers(ctx, [value ?? null]);

/** The two-sample t statistic and degrees of freedom for Excel's three T.TEST types. */
function tStatistic(a: number[], b: number[], type: number): { t: number; df: number } {
	if (type === 1) {
		if (a.length !== b.length) fail(ERR.NA);
		const diffs = a.map((v, i) => v - (b[i] ?? 0));
		const n = diffs.length;
		if (n < 2) fail(ERR.DIV0);
		const sd = S.stdev(diffs, true);
		if (sd === 0) fail(ERR.DIV0);
		return { t: S.mean(diffs) / (sd / Math.sqrt(n)), df: n - 1 };
	}
	if (a.length < 2 || b.length < 2) fail(ERR.DIV0);
	const [n1, n2] = [a.length, b.length];
	const [v1, v2] = [S.variance(a, true), S.variance(b, true)];
	const delta = S.mean(a) - S.mean(b);
	if (type === 2) {
		const pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2);
		if (pooled === 0) fail(ERR.DIV0);
		return { t: delta / Math.sqrt(pooled * (1 / n1 + 1 / n2)), df: n1 + n2 - 2 };
	}
	const [s1, s2] = [v1 / n1, v2 / n2];
	if (s1 + s2 === 0) fail(ERR.DIV0);
	const df = (s1 + s2) ** 2 / (s1 ** 2 / (n1 - 1) + s2 ** 2 / (n2 - 1));
	return { t: delta / Math.sqrt(s1 + s2), df };
}

function studentP(a: number[], b: number[], tails: number, type: number): number {
	if ((tails !== 1 && tails !== 2) || type < 1 || type > 3) fail(ERR.NUM);
	const { t, df } = tStatistic(a, b, type);
	return tails * (1 - tCdf(Math.abs(t), df));
}

function fTestP(a: number[], b: number[]): number {
	if (a.length < 2 || b.length < 2) fail(ERR.DIV0);
	const [v1, v2] = [S.variance(a, true), S.variance(b, true)];
	if (v1 === 0 || v2 === 0) fail(ERR.DIV0);
	const left = fCdf(v1 / v2, a.length - 1, b.length - 1);
	return Math.min(1, 2 * Math.min(left, 1 - left));
}

function chiSquareP(ctx: CallContext, actual: Value, expected: Value): number {
	const [a, e] = [ctx.toMatrix(actual).data, ctx.toMatrix(expected).data];
	const rows = a.length;
	const cols = a[0]?.length ?? 0;
	if (rows !== e.length || cols !== (e[0]?.length ?? 0)) fail(ERR.NA);
	let stat = 0;
	for (let r = 0; r < rows; r++) {
		for (let c = 0; c < cols; c++) {
			const [x, y] = [a[r]?.[c] ?? null, e[r]?.[c] ?? null];
			if (isError(x)) fail(x);
			if (isError(y)) fail(y);
			if (typeof x !== 'number') continue;
			// A blank expected cell counts as 0, which cannot be a divisor.
			if (y === null || y === 0) fail(ERR.DIV0);
			if (typeof y !== 'number') continue;
			stat += (x - y) ** 2 / y;
		}
	}
	const df = rows > 1 && cols > 1 ? (rows - 1) * (cols - 1) : rows * cols - 1;
	if (df < 1) fail(ERR.DIV0);
	return 1 - chiCdf(stat, df);
}

function probability(ctx: CallContext, args: Value[]): number {
	const xs = ctx.toMatrix(args[0] ?? null).flat();
	const ps = ctx.toMatrix(args[1] ?? null).flat();
	if (xs.length !== ps.length) fail(ERR.NA);
	const lower = num(args[2]);
	const upper = args.length > 3 && args[3] !== null ? num(args[3]) : lower;
	let total = 0;
	let hit = 0;
	for (let i = 0; i < xs.length; i++) {
		const [x, p] = [xs[i] ?? null, ps[i] ?? null];
		if (isError(x)) fail(x);
		if (isError(p)) fail(p);
		if (typeof p !== 'number' || p < 0 || p > 1) fail(ERR.NUM);
		total += p;
		if (typeof x === 'number' && x >= lower && x <= upper) hit += p;
	}
	if (Math.abs(total - 1) > 1e-9) fail(ERR.NUM);
	return hit;
}

const withAlias = (names: readonly string[], build: (name: string) => FunctionSpec) =>
	names.map(build);

/** Significance tests on samples (Z.TEST, T.TEST, F.TEST, CHISQ.TEST) and PROB. */
export const HYPOTHESIS_TESTS: FunctionSpec[] = [
	...withAlias(['Z.TEST', 'ZTEST'], (name) =>
		spec(
			name,
			C,
			`${name}(array, x, [sigma])`,
			'The one-tailed P-value of a z-test.',
			2,
			3,
			(args, ctx) => {
				const data = numbersOf(ctx, args[0]);
				if (data.length === 0) fail(ERR.NA);
				const sigma = args.length > 2 && args[2] !== null ? num(args[2]) : S.stdev(data, true);
				if (sigma <= 0) fail(ERR.NUM);
				return 1 - normCdf((S.mean(data) - num(args[1])) / (sigma / Math.sqrt(data.length)));
			},
			ARRAYS,
		),
	),
	...withAlias(['T.TEST', 'TTEST'], (name) =>
		spec(
			name,
			C,
			`${name}(array1, array2, tails, type)`,
			"The probability associated with a Student's t-test.",
			4,
			4,
			(args, ctx) => {
				const tails = int(args[2]);
				const type = int(args[3]);
				if (type === 1) {
					const { xs, ys } = pairedNumbers(ctx, args[0] ?? null, args[1] ?? null);
					return studentP(xs, ys, tails, type);
				}
				return studentP(numbersOf(ctx, args[0]), numbersOf(ctx, args[1]), tails, type);
			},
			ARRAYS,
		),
	),
	...withAlias(['F.TEST', 'FTEST'], (name) =>
		spec(
			name,
			C,
			`${name}(array1, array2)`,
			'The two-tailed probability that the variances of two arrays are not different.',
			2,
			2,
			(args, ctx) => fTestP(numbersOf(ctx, args[0]), numbersOf(ctx, args[1])),
			ARRAYS,
		),
	),
	...withAlias(['CHISQ.TEST', 'CHITEST'], (name) =>
		spec(
			name,
			C,
			`${name}(actual_range, expected_range)`,
			'The chi-squared test of independence.',
			2,
			2,
			(args, ctx) => chiSquareP(ctx, args[0] ?? null, args[1] ?? null),
			ARRAYS,
		),
	),
	spec(
		'PROB',
		C,
		'PROB(x_range, prob_range, lower_limit, [upper_limit])',
		'The probability that values in a range are between two limits.',
		3,
		4,
		(args, ctx) => probability(ctx, args),
		['any', 'any', 'value', 'value'],
	),
];
