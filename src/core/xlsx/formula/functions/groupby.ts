// GROUPBY and PIVOTBY (a documented subset, see viewers/xlsx/docs/features.md):
// one aggregating function (a function name such as SUM, PERCENTOF, or a LAMBDA), grand totals
// only (total depth 0, 1 or -1), no subtotals, no multi-function arrays, flat group output.
import { compareScalars } from '../coerce';
import type { FormulaAst } from '../ast';
import type { CallContext, LazyArg } from '../context';
import type { CellError } from '../../model';
import { ERR, fail, isError, LambdaValue, Matrix, type Scalar, type Value } from '../values';
import { collectNumbers, int } from './helpers';
import type { FunctionSpec } from './types';

type Aggregate = (group: Scalar[], all: Scalar[]) => Scalar;

const column = (values: Scalar[]): Matrix => new Matrix(values.map((v) => [v]));

/** The aggregate named or built by the `function` argument (eta-reduced name or LAMBDA). */
function aggregateOf(arg: LazyArg | undefined, ctx: CallContext): Aggregate {
	if (!arg) return fail(ERR.VALUE);
	const node = arg.node;
	if (node.type === 'name' && !ctx.frame.scope?.has(node.name.toUpperCase())) {
		const name = node.name.toUpperCase();
		if (name === 'PERCENTOF') {
			return (group, all) => {
				const sum = (v: Scalar[]) => collectNumbers(ctx, [column(v)]).reduce((a, b) => a + b, 0);
				const total = sum(all);
				return total === 0 ? ERR.DIV0 : sum(group) / total;
			};
		}
		// Evaluate NAME(_G) with the group bound to a scope variable: any function applies.
		const call: FormulaAst = {
			type: 'call',
			name,
			rawName: node.name,
			args: [{ type: 'name', name: '_G' }],
		};
		return (group) =>
			ctx.toScalar(
				ctx.evaluate(call, new Map([...(ctx.frame.scope ?? []), ['_G', column(group)]])),
			);
	}
	const lambda = arg.get();
	if (isError(lambda)) return fail(lambda);
	if (!(lambda instanceof LambdaValue)) return fail(ERR.VALUE);
	return (group) => ctx.toScalar(ctx.callLambda(lambda, [column(group)]));
}

const optional = (args: LazyArg[], i: number): Value | undefined => {
	const arg = args[i];
	return arg && arg.node.type !== 'missing' ? arg.get() : undefined;
};

function optionalInt(args: LazyArg[], i: number, ctx: CallContext): number | undefined {
	const value = optional(args, i);
	return value === undefined || value === null ? undefined : int(ctx.toScalar(value));
}

function sortSpec(args: LazyArg[], i: number, ctx: CallContext, fallback: number[]): number[] {
	const value = optional(args, i);
	if (value === undefined || value === null) return fallback;
	const list = ctx
		.toMatrix(value)
		.flat()
		.map((v) => (typeof v === 'number' ? Math.trunc(v) : fail(ERR.VALUE)));
	if (list.some((n) => n === 0)) fail(ERR.VALUE);
	return list;
}

function totalDepth(args: LazyArg[], i: number, ctx: CallContext): number {
	const depth = optionalInt(args, i, ctx) ?? 1;
	if (depth !== 0 && depth !== 1 && depth !== -1) fail(ERR.VALUE);
	return depth;
}

function headerMode(args: LazyArg[], i: number, ctx: CallContext, firstRow: Scalar[]): boolean {
	const mode = optionalInt(args, i, ctx);
	if (mode === undefined)
		return firstRow.length > 0 && firstRow.every((v) => typeof v === 'string');
	if (mode === 0) return false;
	if (mode === 1 || mode === 3) return true;
	return fail(ERR.VALUE);
}

const keyOf = (values: Scalar[]): string => JSON.stringify(values);

const plain = (v: Scalar | undefined): Exclude<Scalar, CellError> =>
	v === undefined || isError(v) ? null : v;

function compareBy(keys: number[], rows: Scalar[][]): (a: number, b: number) => number {
	return (a, b) => {
		for (const key of keys) {
			const i = Math.abs(key) - 1;
			const result = compareScalars(plain(rows[a]?.[i]), plain(rows[b]?.[i]));
			if (result !== 0) return key < 0 ? -result : result;
		}
		return a - b;
	};
}

function filtered(rows: number, args: LazyArg[], i: number, ctx: CallContext): number[] {
	const value = optional(args, i);
	const all = Array.from({ length: rows }, (_, r) => r);
	if (value === undefined || value === null) return all;
	const flags = ctx.toMatrix(value).flat();
	if (flags.length !== rows) fail(ERR.VALUE);
	return all.filter((r) => flags[r] === true || (typeof flags[r] === 'number' && flags[r] !== 0));
}

function groupby(args: LazyArg[], ctx: CallContext): Value {
	const keysM = ctx.toMatrix((args[0] as LazyArg).get());
	const valuesM = ctx.toMatrix((args[1] as LazyArg).get());
	if (keysM.rows !== valuesM.rows) fail(ERR.VALUE);
	const aggregate = aggregateOf(args[2], ctx);
	const hasHeaders = headerMode(args, 3, ctx, valuesM.data[0] ?? []);
	const depth = totalDepth(args, 4, ctx);
	const kCols = keysM.cols;
	const vCols = valuesM.cols;
	const sort = sortSpec(
		args,
		5,
		ctx,
		Array.from({ length: kCols }, (_, i) => i + 1),
	);
	const start = hasHeaders ? 1 : 0;
	const rows = filtered(keysM.rows, args, 6, ctx).filter((r) => r >= start);
	if (rows.length === 0) fail(ERR.CALC);
	const groups = new Map<string, number[]>();
	for (const r of rows) {
		const key = keyOf(keysM.data[r] as Scalar[]);
		const found = groups.get(key);
		if (found) found.push(r);
		else groups.set(key, [r]);
	}
	const table = [...groups.values()].map((members) => {
		const key = keysM.data[members[0] as number] as Scalar[];
		const cells = Array.from({ length: vCols }, (_, c) =>
			aggregate(
				members.map((r) => (valuesM.data[r] as Scalar[])[c] ?? null),
				rows.map((r) => (valuesM.data[r] as Scalar[])[c] ?? null),
			),
		);
		return [...key, ...cells];
	});
	for (const key of sort) if (Math.abs(key) > kCols + vCols) fail(ERR.VALUE);
	const order = table.map((_, i) => i).sort(compareBy(sort, table));
	const out: Scalar[][] = order.map((i) => table[i] as Scalar[]);
	if (depth !== 0) {
		const total: Scalar[] = [
			'Total',
			...Array<Scalar>(kCols - 1).fill(''),
			...Array.from({ length: vCols }, (_, c) =>
				aggregate(
					rows.map((r) => (valuesM.data[r] as Scalar[])[c] ?? null),
					rows.map((r) => (valuesM.data[r] as Scalar[])[c] ?? null),
				),
			),
		];
		if (depth === 1) out.push(total);
		else out.unshift(total);
	}
	if (hasHeaders) out.unshift([...(keysM.data[0] as Scalar[]), ...(valuesM.data[0] as Scalar[])]);
	return new Matrix(out);
}

function pivotby(args: LazyArg[], ctx: CallContext): Value {
	const rowsM = ctx.toMatrix((args[0] as LazyArg).get());
	const colsM = ctx.toMatrix((args[1] as LazyArg).get());
	const valuesM = ctx.toMatrix((args[2] as LazyArg).get());
	if (rowsM.rows !== valuesM.rows || colsM.rows !== valuesM.rows || valuesM.cols !== 1)
		fail(ERR.VALUE);
	if (colsM.cols !== 1) fail(ERR.VALUE);
	const aggregate = aggregateOf(args[3], ctx);
	const hasHeaders = headerMode(args, 4, ctx, valuesM.data[0] ?? []);
	const rowDepth = totalDepth(args, 5, ctx);
	const kCols = rowsM.cols;
	const rowSort = sortSpec(
		args,
		6,
		ctx,
		Array.from({ length: kCols }, (_, i) => i + 1),
	);
	const colDepth = totalDepth(args, 7, ctx);
	const colSort = sortSpec(args, 8, ctx, [1]);
	for (const key of rowSort) if (Math.abs(key) > kCols + 1) fail(ERR.VALUE);
	for (const key of colSort) if (Math.abs(key) > 2) fail(ERR.VALUE);
	const start = hasHeaders ? 1 : 0;
	const used = filtered(rowsM.rows, args, 9, ctx).filter((r) => r >= start);
	if (used.length === 0) fail(ERR.CALC);
	const val = (r: number): Scalar => (valuesM.data[r] as Scalar[])[0] ?? null;
	const all = used.map(val);
	const rowKeys = new Map<string, number[]>();
	const colKeys = new Map<string, number[]>();
	for (const r of used) {
		for (const [map, key] of [
			[rowKeys, keyOf(rowsM.data[r] as Scalar[])],
			[colKeys, keyOf(colsM.data[r] as Scalar[])],
		] as const) {
			const found = map.get(key);
			if (found) found.push(r);
			else map.set(key, [r]);
		}
	}
	const agg = (members: number[]): Scalar => aggregate(members.map(val), all);
	const colGroups = [...colKeys.values()].map((members) => ({
		label: (colsM.data[members[0] as number] as Scalar[])[0] ?? null,
		members,
		total: agg(members),
	}));
	const colOrder = colGroups
		.map((_, i) => i)
		.sort(
			compareBy(
				colSort,
				colGroups.map((g) => [g.label, g.total]),
			),
		);
	const rowGroups = [...rowKeys.values()].map((members) => ({
		key: rowsM.data[members[0] as number] as Scalar[],
		members,
		total: agg(members),
	}));
	const rowOrder = rowGroups
		.map((_, i) => i)
		.sort(
			compareBy(
				rowSort,
				rowGroups.map((g) => [...g.key, g.total]),
			),
		);
	const columns = colOrder.map((i) => colGroups[i] as (typeof colGroups)[number]);
	const lines: Scalar[][] = rowOrder.map((i) => {
		const group = rowGroups[i] as (typeof rowGroups)[number];
		const cells = columns.map((col) => {
			const members = group.members.filter((r) => col.members.includes(r));
			return members.length ? agg(members) : '';
		});
		return [
			...group.key,
			...(colDepth === 1
				? [...cells, group.total]
				: colDepth === -1
					? [group.total, ...cells]
					: cells),
		];
	});
	const grandRow: Scalar[] = [
		'Total',
		...Array<Scalar>(kCols - 1).fill(''),
		...columns.map((col) => col.total),
		...(colDepth === 0 ? [] : [agg(used)]),
	];
	if (colDepth === -1)
		grandRow.splice(kCols, columns.length + 1, agg(used), ...columns.map((c) => c.total));
	if (rowDepth === 1) lines.push(grandRow);
	else if (rowDepth === -1) lines.unshift(grandRow);
	const header: Scalar[] = [
		...(hasHeaders
			? [(rowsM.data[0] as Scalar[]).join(' '), ...Array<Scalar>(kCols - 1).fill('')]
			: Array<Scalar>(kCols).fill('')),
		...(colDepth === -1 ? ['Total'] : []),
		...columns.map((col) => col.label),
		...(colDepth === 1 ? ['Total'] : []),
	];
	return new Matrix([header, ...lines]);
}

const lazy = (
	name: string,
	syntax: string,
	description: string,
	minArgs: number,
	maxArgs: number,
	fn: (args: LazyArg[], ctx: CallContext) => Value,
): FunctionSpec => ({
	name,
	category: 'Lookup & Reference',
	syntax,
	description,
	minArgs,
	maxArgs,
	lazy: fn,
});

export const GROUPBY_FUNCTIONS: FunctionSpec[] = [
	lazy(
		'GROUPBY',
		'GROUPBY(row_fields, values, function, [field_headers], [total_depth], [sort_order], [filter_array], [field_relationship])',
		'Groups rows by the row fields and aggregates the values.',
		3,
		8,
		groupby,
	),
	lazy(
		'PIVOTBY',
		'PIVOTBY(row_fields, col_fields, values, function, [field_headers], [row_total_depth], [row_sort_order], [col_total_depth], [col_sort_order], [filter_array], [relative_to])',
		'Groups rows and columns by fields and aggregates the values into a table.',
		4,
		11,
		pivotby,
	),
];
