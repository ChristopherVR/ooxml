import type { CallContext } from '../context.js';
import { ERR, fail, isError, Matrix, RefValue, type Scalar, type Value } from '../values.js';
import { bool, int, num, optNum, scalar, spec } from './helpers.js';
import { findExact, findSorted, line, shape, vectorOf, xsearch } from './lookup-core.js';
import { REFERENCE_FUNCTIONS } from './reference.js';
import type { FunctionSpec } from './types.js';

const C = 'Lookup & Reference';

/** The cell (or matrix element) at a 0-based position of a ref or array, as a reference when possible. */
export function cellOf(value: Value, row: number, col: number): Value {
	if (value instanceof RefValue) {
		const area = value.areas[0];
		if (!area) fail(ERR.REF);
		const r = area.range.start.row + row;
		const c = area.range.start.col + col;
		return new RefValue([
			{ sheet: area.sheet, range: { start: { row: r, col: c }, end: { row: r, col: c } } },
		]);
	}
	if (value instanceof Matrix) return value.get(row, col);
	return value as Scalar;
}

/** Row or column `index` (0-based) of a ref or array, as a ref or matrix. */
export function sliceOf(value: Value, axis: 'row' | 'col', index: number): Value {
	const { rows } = shape(value);
	if (value instanceof RefValue) {
		const area = value.areas[0];
		if (!area) fail(ERR.REF);
		const { start, end } = area.range;
		const range =
			axis === 'row'
				? {
						start: { row: start.row + index, col: start.col },
						end: { row: start.row + index, col: end.col },
					}
				: {
						start: { row: start.row, col: start.col + index },
						end: { row: end.row, col: start.col + index },
					};
		return new RefValue([{ sheet: area.sheet, range }]);
	}
	const m = value instanceof Matrix ? value : new Matrix([[value as Scalar]]);
	if (axis === 'row') return new Matrix([m.data[index] ?? []]);
	return new Matrix(Array.from({ length: rows }, (_, r) => [m.get(r, index)]));
}

function lookupValue(args: Value[]): Scalar {
	const v = scalar(args[0]);
	if (isError(v)) fail(v);
	return v;
}

function vhlookup(args: Value[], ctx: CallContext, vertical: boolean): Value {
	const lookup = lookupValue(args);
	const table = args[1] ?? null;
	const index = int(args[2]);
	const approximate = args.length < 4 ? true : bool(args[3]);
	const { rows, cols } = shape(table);
	if (index < 1) fail(ERR.VALUE);
	if (index > (vertical ? cols : rows)) fail(ERR.REF);
	const keys = line(ctx, table, vertical ? 'col' : 'row', 0);
	const at = approximate ? findSorted(keys, lookup) : findExact(keys, lookup, true);
	if (at < 0) fail(ERR.NA);
	return vertical ? cellOf(table, at, index - 1) : cellOf(table, index - 1, at);
}

function match(args: Value[], ctx: CallContext): number {
	const lookup = lookupValue(args);
	const type = Math.sign(optNum(args, 2, 1));
	const { vector } = vectorOf(ctx, args[1] ?? null);
	const at = type === 0 ? findExact(vector, lookup, true) : findSorted(vector, lookup, type < 0);
	return at < 0 ? fail(ERR.NA) : at + 1;
}

function lookup(args: Value[], ctx: CallContext): Value {
	const value = lookupValue(args);
	const source = args[1] ?? null;
	const { rows, cols } = shape(source);
	if (args.length > 2) {
		const keys = vectorOf(ctx, source).vector;
		const at = findSorted(keys, value);
		if (at < 0) fail(ERR.NA);
		const result = args[2] ?? null;
		const rs = shape(result);
		return rs.cols === 1 ? cellOf(result, at, 0) : cellOf(result, 0, at);
	}
	const vertical = rows >= cols;
	const keys = line(ctx, source, vertical ? 'col' : 'row', 0);
	const at = findSorted(keys, value);
	if (at < 0) fail(ERR.NA);
	return vertical ? cellOf(source, at, cols - 1) : cellOf(source, rows - 1, at);
}

function index(args: Value[]): Value {
	const source = args[0] ?? null;
	if (source instanceof RefValue && source.areas.length > 1) {
		const area = Math.trunc(optNum(args, 3, 1));
		const chosen = source.areas[area - 1];
		if (!chosen) fail(ERR.REF);
		return index([new RefValue([chosen]), ...args.slice(1, 3)]);
	}
	const { rows, cols } = shape(source);
	let r = args.length > 1 && args[1] !== null ? int(args[1]) : 0;
	let c = args.length > 2 && args[2] !== null ? int(args[2]) : 0;
	if (args.length <= 2 && rows === 1 && cols > 1) {
		c = r;
		r = 1;
	}
	if (r < 0 || c < 0 || r > rows || c > cols) fail(ERR.REF);
	if (r === 0 && c === 0) return source;
	if (r === 0) return cols === 1 ? source : sliceOf(source, 'col', c - 1);
	if (c === 0)
		return rows === 1 || cols === 1 ? cellOf(source, r - 1, 0) : sliceOf(source, 'row', r - 1);
	return cellOf(source, r - 1, c - 1);
}

function xlookup(args: Value[], ctx: CallContext): Value {
	const value = ctx.toScalar(args[0] ?? null);
	if (isError(value)) return value;
	const { vector, axis } = vectorOf(ctx, args[1] ?? null);
	const result = args[2] ?? null;
	const resultShape = shape(result);
	const matchMode = Math.trunc(
		args.length > 4 && args[4] !== null ? num(ctx.toScalar(args[4] ?? null)) : 0,
	);
	const searchMode = Math.trunc(
		args.length > 5 && args[5] !== null ? num(ctx.toScalar(args[5] ?? null)) : 1,
	);
	const lookupLength = shape(args[1] ?? null)[axis === 'col' ? 'rows' : 'cols'];
	if ((axis === 'col' ? resultShape.rows : resultShape.cols) !== lookupLength) fail(ERR.VALUE);
	const at = xsearch(vector, value, matchMode, searchMode);
	if (at < 0) {
		if (args.length > 3 && args[3] !== null) return args[3] ?? ERR.NA;
		return ERR.NA;
	}
	if (axis === 'col')
		return resultShape.cols === 1 ? cellOf(result, at, 0) : sliceOf(result, 'row', at);
	return resultShape.rows === 1 ? cellOf(result, 0, at) : sliceOf(result, 'col', at);
}

export const LOOKUP_FUNCTIONS: FunctionSpec[] = [
	...REFERENCE_FUNCTIONS,
	spec(
		'VLOOKUP',
		C,
		'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
		'Looks up a value in the first column of a table.',
		3,
		4,
		(args, ctx) => vhlookup(args, ctx, true),
		['value', 'any', 'value', 'value'],
	),
	spec(
		'HLOOKUP',
		C,
		'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])',
		'Looks up a value in the first row of a table.',
		3,
		4,
		(args, ctx) => vhlookup(args, ctx, false),
		['value', 'any', 'value', 'value'],
	),
	spec(
		'LOOKUP',
		C,
		'LOOKUP(lookup_value, lookup_vector, [result_vector])',
		'Looks up a value in a sorted vector or array.',
		2,
		3,
		(args, ctx) => lookup(args, ctx),
		['value', 'any', 'any'],
	),
	spec(
		'MATCH',
		C,
		'MATCH(lookup_value, lookup_array, [match_type])',
		'The position of a value in a range.',
		2,
		3,
		(args, ctx) => match(args, ctx),
		['value', 'any', 'value'],
	),
	spec(
		'XMATCH',
		C,
		'XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])',
		'The position of a value with flexible matching.',
		2,
		4,
		(args, ctx) => {
			const value = lookupValue(args);
			const { vector } = vectorOf(ctx, args[1] ?? null);
			const at = xsearch(
				vector,
				value,
				Math.trunc(optNum(args, 2, 0)),
				Math.trunc(optNum(args, 3, 1)),
			);
			return at < 0 ? fail(ERR.NA) : at + 1;
		},
		['value', 'any', 'value', 'value'],
	),
	spec(
		'XLOOKUP',
		C,
		'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])',
		'Searches a range and returns the matching item from another.',
		3,
		6,
		(args, ctx) => xlookup(args, ctx),
		['value', 'any'],
	),
	spec(
		'INDEX',
		C,
		'INDEX(array, row_num, [column_num], [area_num])',
		'The value or reference at a position.',
		1,
		4,
		(args) => index(args),
		['any', 'value', 'value', 'value'],
	),
	spec(
		'CHOOSE',
		C,
		'CHOOSE(index_num, value1, [value2], ...)',
		'Chooses a value from a list by index.',
		2,
		255,
		(args) => {
			const i = int(args[0]);
			if (i < 1 || i >= args.length) fail(ERR.VALUE);
			return args[i] ?? null;
		},
		['value', 'any'],
	),
];
