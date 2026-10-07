import type { CallContext } from '../context';
import { ERR, fail, isError, type Scalar, type Value } from '../values';
import { makeCriteria, scalar, spec } from './helpers';
import * as S from './stats-core';
import type { FunctionSpec } from './types';

const C = 'Database';

/** A list with a header row; every following row is a record. */
interface Table {
	headers: string[];
	records: Scalar[][];
}

const label = (value: Scalar): string =>
	value === null ? '' : typeof value === 'string' ? value : String(value);

function tableOf(ctx: CallContext, value: Value | undefined): Table {
	const [head, ...records] = ctx.toMatrix(value ?? null).data;
	if (!head || records.length === 0) fail(ERR.VALUE);
	return { headers: head.map((v) => label(v).toLowerCase()), records };
}

/** The zero-based column a `field` argument names: a 1-based index or a header label. */
function fieldIndex(table: Table, field: Value | undefined): number {
	const value = scalar(field);
	if (isError(value)) fail(value);
	if (typeof value === 'number') {
		const index = Math.trunc(value) - 1;
		if (index < 0 || index >= table.headers.length) fail(ERR.VALUE);
		return index;
	}
	const index = table.headers.indexOf(label(value).toLowerCase());
	if (index < 0) fail(ERR.VALUE);
	return index;
}

/**
 * The record filter of a criteria range: its first row names fields, each later row is an
 * alternative (OR) whose non-blank cells must all match (AND). A blank criteria cell is ignored.
 */
function recordFilter(table: Table, criteria: Scalar[][]): (record: readonly Scalar[]) => boolean {
	const [head, ...rows] = criteria;
	if (!head || rows.length === 0) fail(ERR.VALUE);
	const columns = head.map((name) => {
		if (name === null || name === '') return -1;
		return table.headers.indexOf(label(name).toLowerCase());
	});
	const alternatives = rows.map((row) =>
		row.flatMap((cell, c) => {
			const column = columns[c] ?? -1;
			if (cell === null || cell === '') return [];
			// A criteria column naming no database field can never match.
			if (column < 0) return [() => false];
			const matches = makeCriteria(cell);
			return [(record: readonly Scalar[]) => matches(record[column] ?? null)];
		}),
	);
	return (record) => alternatives.some((tests) => tests.every((test) => test(record)));
}

/** The values of `field` in the records matching `criteria`. */
function selected(ctx: CallContext, args: Value[]): Scalar[] {
	const table = tableOf(ctx, args[0]);
	const column = fieldIndex(table, args[1]);
	const matches = recordFilter(table, ctx.toMatrix(args[2] ?? null).data);
	return table.records.filter(matches).map((record) => record[column] ?? null);
}

const numbers = (values: Scalar[]): number[] => {
	for (const v of values) if (isError(v)) fail(v);
	return values.filter((v): v is number => typeof v === 'number');
};

const dbFn = (name: string, description: string, fn: (values: Scalar[]) => Scalar): FunctionSpec =>
	spec(
		name,
		C,
		`${name}(database, field, criteria)`,
		description,
		3,
		3,
		(args, ctx) => fn(selected(ctx, args)),
		['any', 'any', 'any'],
	);

/** Database functions: aggregate one column of the records that match a criteria range. */
export const DATABASE_FUNCTIONS: FunctionSpec[] = [
	dbFn('DSUM', 'The sum of the matching records in a column.', (v) => S.sum(numbers(v))),
	dbFn('DCOUNT', 'The count of numbers in the matching records.', (v) => numbers(v).length),
	dbFn(
		'DCOUNTA',
		'The count of non-blank cells in the matching records.',
		(v) => v.filter((x) => x !== null && x !== '').length,
	),
	dbFn('DAVERAGE', 'The mean of the matching records in a column.', (v) => S.mean(numbers(v))),
	dbFn('DMAX', 'The largest number among the matching records.', (v) => {
		const n = numbers(v);
		return n.length === 0 ? 0 : S.extreme(n, true);
	}),
	dbFn('DMIN', 'The smallest number among the matching records.', (v) => {
		const n = numbers(v);
		return n.length === 0 ? 0 : S.extreme(n, false);
	}),
	dbFn('DPRODUCT', 'The product of the matching records in a column.', (v) => {
		const n = numbers(v);
		return n.length === 0 ? 0 : n.reduce((a, b) => a * b, 1);
	}),
	dbFn('DGET', 'The single value of the one matching record.', (v) => {
		if (v.length === 0) fail(ERR.VALUE);
		if (v.length > 1) fail(ERR.NUM);
		return v[0] ?? null;
	}),
	dbFn('DSTDEV', 'The sample standard deviation of the matching records.', (v) =>
		S.stdev(numbers(v), true),
	),
	dbFn('DSTDEVP', 'The population standard deviation of the matching records.', (v) =>
		S.stdev(numbers(v), false),
	),
	dbFn('DVAR', 'The sample variance of the matching records.', (v) => S.variance(numbers(v), true)),
	dbFn('DVARP', 'The population variance of the matching records.', (v) =>
		S.variance(numbers(v), false),
	),
];
