import { compareScalars, toBool } from '../coerce.js';
import type { CallContext } from '../context.js';
import { ERR, fail, isError, Matrix, type Scalar, type Value } from '../values.js';
import { ARRAY_SHAPE_FUNCTIONS } from './array-shape.js';
import { num, spec } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Lookup & Reference';

/** Sort order for SORT/SORTBY: numbers, text, logicals, errors; blanks always last. */
export function sortCompare(a: Scalar, b: Scalar, descending: boolean): number {
	if (a === null || b === null) return a === b ? 0 : a === null ? 1 : -1;
	const ea = isError(a);
	const eb = isError(b);
	let c: number;
	if (ea || eb) c = ea && eb ? 0 : ea ? 1 : -1;
	else c = compareScalars(a, b);
	return descending ? -c : c;
}

const optInt = (ctx: CallContext, args: Value[], i: number, fallback: number): number =>
	i < args.length && args[i] !== null ? Math.trunc(num(ctx.toScalar(args[i] ?? null))) : fallback;

function transpose(m: Matrix): Matrix {
	return Matrix.build(m.cols, m.rows, (r, c) => m.get(c, r));
}

function filter(args: Value[], ctx: CallContext): Value {
	const array = ctx.toMatrix(args[0] ?? null);
	const include = ctx.toMatrix(args[1] ?? null);
	const truthy = (v: Scalar): boolean => {
		if (isError(v)) fail(v);
		return v !== null && toBool(v);
	};
	let out: Matrix;
	if (include.cols === 1 && include.rows === array.rows) {
		out = new Matrix(array.data.filter((_, r) => truthy(include.get(r, 0))));
	} else if (include.rows === 1 && include.cols === array.cols) {
		const keep = include.data[0]?.map((v) => truthy(v)) ?? [];
		out = new Matrix(array.data.map((row) => row.filter((_, c) => keep[c])));
	} else {
		return ERR.VALUE;
	}
	if (out.rows === 0 || out.cols === 0) {
		return args.length > 2 ? (args[2] ?? ERR.CALC) : ERR.CALC;
	}
	return out;
}

function sortRows(rows: Scalar[][], keys: { values: Scalar[]; descending: boolean }[]): Scalar[][] {
	const order = rows.map((_, i) => i);
	order.sort((x, y) => {
		for (const key of keys) {
			const c = sortCompare(key.values[x] ?? null, key.values[y] ?? null, key.descending);
			if (c !== 0) return c;
		}
		return x - y;
	});
	return order.map((i) => rows[i] ?? []);
}

function sort(args: Value[], ctx: CallContext): Matrix {
	let m = ctx.toMatrix(args[0] ?? null);
	const byCol = args.length > 3 && args[3] !== null && toBool(ctx.toScalar(args[3] ?? null));
	if (byCol) m = transpose(m);
	const indexes =
		args.length > 1 && args[1] !== null
			? ctx
					.toMatrix(args[1] ?? null)
					.flat()
					.map((v) => Math.trunc(num(v)))
			: [1];
	const orders =
		args.length > 2 && args[2] !== null
			? ctx
					.toMatrix(args[2] ?? null)
					.flat()
					.map((v) => num(v))
			: [1];
	const keys = indexes.map((index, k) => {
		if (index < 1 || index > m.cols) fail(ERR.VALUE);
		const order = orders[Math.min(k, orders.length - 1)] ?? 1;
		if (order !== 1 && order !== -1) fail(ERR.VALUE);
		return { values: m.data.map((row) => row[index - 1] ?? null), descending: order === -1 };
	});
	const sorted = new Matrix(sortRows(m.data, keys));
	return byCol ? transpose(sorted) : sorted;
}

function sortBy(args: Value[], ctx: CallContext): Matrix {
	const m = ctx.toMatrix(args[0] ?? null);
	const keys: { values: Scalar[]; descending: boolean }[] = [];
	let byCol: boolean | undefined;
	for (let i = 1; i < args.length; i += 2) {
		const by = ctx.toMatrix(args[i] ?? null);
		const order =
			i + 1 < args.length && args[i + 1] !== null ? num(ctx.toScalar(args[i + 1] ?? null)) : 1;
		if (order !== 1 && order !== -1) fail(ERR.VALUE);
		const column = by.cols === 1 && by.rows === m.rows;
		const row = by.rows === 1 && by.cols === m.cols;
		if (!column && !row) fail(ERR.VALUE);
		const thisByCol = !column;
		if (byCol !== undefined && byCol !== thisByCol) fail(ERR.VALUE);
		byCol = thisByCol;
		keys.push({ values: by.flat(), descending: order === -1 });
	}
	if (byCol) return transpose(new Matrix(sortRows(transpose(m).data, keys)));
	return new Matrix(sortRows(m.data, keys));
}

function unique(args: Value[], ctx: CallContext): Value {
	let m = ctx.toMatrix(args[0] ?? null);
	const byCol = args.length > 1 && args[1] !== null && toBool(ctx.toScalar(args[1] ?? null));
	const once = args.length > 2 && args[2] !== null && toBool(ctx.toScalar(args[2] ?? null));
	if (byCol) m = transpose(m);
	const keyOf = (row: Scalar[]): string =>
		JSON.stringify(
			row.map((v) =>
				isError(v) ? `#${v.error}` : typeof v === 'string' ? `s${v.toLowerCase()}` : v,
			),
		);
	const counts = new Map<string, { row: Scalar[]; count: number }>();
	for (const row of m.data) {
		const key = keyOf(row);
		const entry = counts.get(key);
		if (entry) entry.count++;
		else counts.set(key, { row, count: 1 });
	}
	const rows = [...counts.values()].filter((e) => !once || e.count === 1).map((e) => e.row);
	if (rows.length === 0) return ERR.CALC;
	const out = new Matrix(rows);
	return byCol ? transpose(out) : out;
}

export const ARRAY_FUNCTIONS: FunctionSpec[] = [
	...ARRAY_SHAPE_FUNCTIONS,
	spec(
		'FILTER',
		C,
		'FILTER(array, include, [if_empty])',
		'Filters an array by a Boolean array.',
		2,
		3,
		(args, ctx) => filter(args, ctx),
		['any'],
	),
	spec(
		'SORT',
		C,
		'SORT(array, [sort_index], [sort_order], [by_col])',
		'Sorts the contents of an array.',
		1,
		4,
		(args, ctx) => sort(args, ctx),
		['any'],
	),
	spec(
		'SORTBY',
		C,
		'SORTBY(array, by_array1, [sort_order1], ...)',
		'Sorts an array by other arrays.',
		2,
		255,
		(args, ctx) => sortBy(args, ctx),
		['any'],
	),
	spec(
		'UNIQUE',
		C,
		'UNIQUE(array, [by_col], [exactly_once])',
		'The unique rows or columns of an array.',
		1,
		3,
		(args, ctx) => unique(args, ctx),
		['any'],
	),
	spec(
		'TRANSPOSE',
		C,
		'TRANSPOSE(array)',
		'Swaps the rows and columns of an array.',
		1,
		1,
		(args, ctx) => transpose(ctx.toMatrix(args[0] ?? null).map((v) => (v === null ? 0 : v))),
		['any'],
	),
	spec(
		'SEQUENCE',
		C,
		'SEQUENCE(rows, [columns], [start], [step])',
		'An array of sequential numbers.',
		1,
		4,
		(args, ctx) => {
			const rows = optInt(ctx, args, 0, 1);
			const cols = optInt(ctx, args, 1, 1);
			const start = args.length > 2 && args[2] !== null ? num(ctx.toScalar(args[2] ?? null)) : 1;
			const step = args.length > 3 && args[3] !== null ? num(ctx.toScalar(args[3] ?? null)) : 1;
			if (rows < 1 || cols < 1) return ERR.CALC;
			if (rows * cols > 1_048_576 * 16) fail(ERR.NUM);
			return Matrix.build(rows, cols, (r, c) => start + (r * cols + c) * step);
		},
		['any'],
	),
	spec(
		'RANDARRAY',
		C,
		'RANDARRAY([rows], [columns], [min], [max], [integer])',
		'An array of random numbers.',
		0,
		5,
		(args, ctx) => {
			const rows = optInt(ctx, args, 0, 1);
			const cols = optInt(ctx, args, 1, 1);
			const min = args.length > 2 && args[2] !== null ? num(ctx.toScalar(args[2] ?? null)) : 0;
			const max = args.length > 3 && args[3] !== null ? num(ctx.toScalar(args[3] ?? null)) : 1;
			const integer = args.length > 4 && args[4] !== null && toBool(ctx.toScalar(args[4] ?? null));
			if (rows < 1 || cols < 1) return ERR.CALC;
			if (min > max) fail(ERR.VALUE);
			const random = (): number => ctx.frame.host.random();
			return Matrix.build(rows, cols, () =>
				integer
					? Math.ceil(min) + Math.floor(random() * (Math.floor(max) - Math.ceil(min) + 1))
					: min + random() * (max - min),
			);
		},
		['any'],
		true,
	),
];
