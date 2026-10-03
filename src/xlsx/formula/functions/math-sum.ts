import { toNumber } from '../coerce.js';
import type { CallContext } from '../context.js';
import {
	ERR,
	ErrorSignal,
	fail,
	isError,
	Matrix,
	RefValue,
	type Scalar,
	type Value,
} from '../values.js';
import { criteriaPairs, liftCriteria, matchingValues } from './criteria.js';
import { collectNumbers, int, spec } from './helpers.js';
import * as S from './stats-core.js';
import type { FunctionSpec } from './types.js';

const C = 'Math & Trig';

const numbersOf = (values: Scalar[]): number[] =>
	values.filter((v): v is number => typeof v === 'number');

function sumProduct(ctx: CallContext, args: Value[]): number {
	const matrices = args.map((a) => ctx.toMatrix(a));
	const first = matrices[0];
	if (!first) fail(ERR.VALUE);
	for (const m of matrices) if (m.rows !== first.rows || m.cols !== first.cols) fail(ERR.VALUE);
	let total = 0;
	for (let r = 0; r < first.rows; r++) {
		for (let c = 0; c < first.cols; c++) {
			let product = 1;
			for (const m of matrices) {
				const v = m.get(r, c);
				if (isError(v)) fail(v);
				product *= typeof v === 'number' ? v : 0;
			}
			total += product;
		}
	}
	return total;
}

function pairwise(ctx: CallContext, args: Value[], fn: (x: number, y: number) => number): number {
	const xs = ctx.toMatrix(args[0] ?? null).flat();
	const ys = ctx.toMatrix(args[1] ?? null).flat();
	if (xs.length !== ys.length) fail(ERR.NA);
	let total = 0;
	for (let i = 0; i < xs.length; i++) {
		const x = xs[i] ?? null;
		const y = ys[i] ?? null;
		if (isError(x)) fail(x);
		if (isError(y)) fail(y);
		if (typeof x === 'number' && typeof y === 'number') total += fn(x, y);
	}
	return total;
}

const SUBTOTAL_RE = /\b(?:_xlfn\.)?(?:SUBTOTAL|AGGREGATE)\s*\(/i;

/** Values for SUBTOTAL / AGGREGATE: skips nested subtotals and, optionally, hidden rows and errors. */
function subtotalValues(
	ctx: CallContext,
	args: Value[],
	skipHidden: boolean,
	skipErrors: boolean,
	skipNested = true,
): Scalar[] {
	const out: Scalar[] = [];
	const host = ctx.frame.host;
	for (const arg of args) {
		if (!(arg instanceof RefValue)) {
			ctx.forEach(arg, (v) => {
				if (!(skipErrors && isError(v))) out.push(v);
			});
			continue;
		}
		for (const area of arg.areas) {
			const sheet = ctx.workbook.sheets[area.sheet];
			host.forEachStored(area.sheet, area.range, (value, row, col) => {
				if (skipHidden && sheet?.rowInfo.get(row)?.hidden) return;
				if (skipNested && SUBTOTAL_RE.test(host.cellFormula(area.sheet, row, col) ?? '')) return;
				if (skipErrors && isError(value)) return;
				out.push(value);
			});
		}
	}
	return out;
}

/** The aggregations behind SUBTOTAL (1-11) and AGGREGATE (1-19). */
export function aggregate(code: number, values: Scalar[], k?: number): number {
	// COUNT and COUNTA count around errors; every other aggregation propagates the first one.
	if (code === 2) return numbersOf(values).length;
	if (code === 3) return values.filter((v) => v !== null).length;
	for (const v of values) if (isError(v)) fail(v);
	const nums = numbersOf(values);
	switch (code) {
		case 1:
			return S.mean(nums);
		case 4:
			return S.extreme(nums, true);
		case 5:
			return S.extreme(nums, false);
		case 6:
			return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
		case 7:
			return S.stdev(nums, true);
		case 8:
			return S.stdev(nums, false);
		case 9:
			return S.sum(nums);
		case 10:
			return S.variance(nums, true);
		case 11:
			return S.variance(nums, false);
		case 12:
			return S.median(nums);
		case 13:
			return S.modes(nums)[0] as number;
		case 14:
			return S.kth(nums, k ?? 1, true);
		case 15:
			return S.kth(nums, k ?? 1, false);
		case 16:
			return S.percentileInc(nums, k ?? 0);
		case 17:
			return S.quartile(nums, k ?? 0, false);
		case 18:
			return S.percentileExc(nums, k ?? 0);
		case 19:
			return S.quartile(nums, k ?? 0, true);
		default:
			return fail(ERR.VALUE);
	}
}

export const SUM_FUNCTIONS: FunctionSpec[] = [
	spec(
		'SUM',
		C,
		'SUM(number1, [number2], ...)',
		'Adds its arguments.',
		0,
		255,
		(args, ctx) => S.sum(collectNumbers(ctx, args)),
		['any'],
	),
	spec(
		'PRODUCT',
		C,
		'PRODUCT(number1, [number2], ...)',
		'Multiplies its arguments.',
		1,
		255,
		(args, ctx) => {
			// With no numbers at all PRODUCT is 0, not the empty product 1.
			const nums = collectNumbers(ctx, args);
			return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
		},
		['any'],
	),
	spec(
		'SUMSQ',
		C,
		'SUMSQ(number1, [number2], ...)',
		'The sum of the squares of the arguments.',
		1,
		255,
		(args, ctx) => collectNumbers(ctx, args).reduce((a, b) => a + b * b, 0),
		['any'],
	),
	spec(
		'SUMPRODUCT',
		C,
		'SUMPRODUCT(array1, [array2], ...)',
		'The sum of the products of corresponding array elements.',
		1,
		255,
		(args, ctx) => sumProduct(ctx, args),
		['any'],
	),
	spec(
		'SUMX2MY2',
		C,
		'SUMX2MY2(array_x, array_y)',
		'Sum of the differences of squares.',
		2,
		2,
		(args, ctx) => pairwise(ctx, args, (x, y) => x * x - y * y),
		['any'],
	),
	spec(
		'SUMX2PY2',
		C,
		'SUMX2PY2(array_x, array_y)',
		'Sum of the sums of squares.',
		2,
		2,
		(args, ctx) => pairwise(ctx, args, (x, y) => x * x + y * y),
		['any'],
	),
	spec(
		'SUMXMY2',
		C,
		'SUMXMY2(array_x, array_y)',
		'Sum of the squares of differences.',
		2,
		2,
		(args, ctx) => pairwise(ctx, args, (x, y) => (x - y) * (x - y)),
		['any'],
	),
	spec(
		'SUMIF',
		C,
		'SUMIF(range, criteria, [sum_range])',
		'Adds the cells that meet a criterion.',
		2,
		3,
		(args, ctx) => {
			const pairs = [{ range: args[0] ?? null, criteria: args[1] ?? null }];
			return S.sum(numbersOf(matchingValues(ctx, pairs, args[2] ?? undefined, false)));
		},
		['any', 'value', 'any'],
	),
	spec(
		'SUMIFS',
		C,
		'SUMIFS(sum_range, criteria_range1, criteria1, ...)',
		'Adds the cells that meet several criteria.',
		3,
		255,
		(args, ctx) =>
			liftCriteria(ctx, criteriaPairs(args, 1), (pairs) =>
				S.sum(numbersOf(matchingValues(ctx, pairs, args[0] ?? null, true))),
			),
		['any'],
	),
	spec(
		'SUBTOTAL',
		C,
		'SUBTOTAL(function_num, ref1, ...)',
		'A subtotal (1-11 include hidden rows, 101-111 ignore them).',
		2,
		255,
		(args, ctx) => {
			const code = int(args[0]);
			const base = code > 100 ? code - 100 : code;
			if (base < 1 || base > 11) fail(ERR.VALUE);
			return aggregate(base, subtotalValues(ctx, args.slice(1), code > 100, false));
		},
		['value', 'any'],
	),
	spec(
		'AGGREGATE',
		C,
		'AGGREGATE(function_num, options, ref1, ...)',
		'An aggregate that can ignore hidden rows and errors.',
		3,
		255,
		(args, ctx) => {
			const code = int(args[0]);
			const options = int(args[1]);
			if (code < 1 || code > 19 || options < 0 || options > 7) fail(ERR.VALUE);
			const skipNested = options <= 3;
			const skipHidden = options === 1 || options === 3 || options === 5 || options === 7;
			const skipErrors = options === 2 || options === 3 || options === 6 || options === 7;
			if (code >= 14) {
				// The array form: one array or reference plus k, which lifts like a value argument.
				if (args.length !== 4) fail(ERR.VALUE);
				const values = subtotalValues(ctx, [args[2] ?? null], skipHidden, skipErrors, skipNested);
				const k = args[3] ?? null;
				const one = (kv: Scalar): Scalar => {
					if (isError(kv)) return kv;
					try {
						return aggregate(code, values, toNumber(kv));
					} catch (e) {
						if (e instanceof ErrorSignal) return e.value;
						throw e;
					}
				};
				if (k instanceof RefValue && k.isCell()) return one(ctx.toScalar(k));
				if (k instanceof RefValue || k instanceof Matrix) return ctx.toMatrix(k).map(one);
				return aggregate(code, values, toNumber(ctx.toScalar(k)));
			}
			// The reference form takes references only.
			if (args.slice(2).some((a) => !(a instanceof RefValue))) fail(ERR.VALUE);
			return aggregate(
				code,
				subtotalValues(ctx, args.slice(2), skipHidden, skipErrors, skipNested),
			);
		},
		['value', 'value', 'any'],
	),
];
