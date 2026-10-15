// PERCENTOF and TRIMRANGE.
import { ERR, fail, Matrix, type Scalar } from '../values';
import { collectNumbers, spec } from './helpers';
import type { FunctionSpec } from './types';

const sum = (values: number[]): number => values.reduce((a, b) => a + b, 0);

/** Leading and trailing blank lines to drop; `mode` 0 none, 1 leading, 2 trailing, 3 both. */
function trimmed(blank: boolean[], mode: number): [number, number] {
	let start = 0;
	let end = blank.length;
	if (mode === 1 || mode === 3) while (start < end && blank[start]) start++;
	if (mode === 2 || mode === 3) while (end > start && blank[end - 1]) end--;
	return [start, end];
}

export const MODERN_FUNCTIONS: FunctionSpec[] = [
	spec(
		'PERCENTOF',
		'Math & Trig',
		'PERCENTOF(data_subset, data_all)',
		'The share a subset of values is of all values.',
		2,
		2,
		(args, ctx) => {
			const total = sum(collectNumbers(ctx, [args[1] ?? null]));
			if (total === 0) fail(ERR.DIV0);
			return sum(collectNumbers(ctx, [args[0] ?? null])) / total;
		},
		['any'],
	),
	spec(
		'TRIMRANGE',
		'Lookup & Reference',
		'TRIMRANGE(range, [trim_rows], [trim_cols])',
		'Removes blank rows and columns from the edges of a range.',
		1,
		3,
		(args, ctx) => {
			const mode = (i: number): number => {
				const v = args[i];
				if (v === undefined || v === null) return 3;
				const n = Math.trunc(ctx.toScalar(v) as number);
				if (typeof ctx.toScalar(v) !== 'number' || n < 0 || n > 3) fail(ERR.VALUE);
				return n;
			};
			const data = ctx.toMatrix(args[0] ?? null).data;
			const cols = data[0]?.length ?? 0;
			const [r0, r1] = trimmed(
				data.map((row) => row.every((v) => v === null)),
				mode(1),
			);
			const [c0, c1] = trimmed(
				Array.from({ length: cols }, (_, c) => data.every((row) => row[c] === null)),
				mode(2),
			);
			if (r1 <= r0 || c1 <= c0) fail(ERR.CALC);
			return new Matrix(data.slice(r0, r1).map((row): Scalar[] => row.slice(c0, c1)));
		},
		['any'],
	),
];
