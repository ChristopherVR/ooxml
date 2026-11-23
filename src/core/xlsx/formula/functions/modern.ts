// PERCENTOF and TRIMRANGE.
import { ERR, fail, Matrix, RefValue, type Scalar } from '../values';
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

/** Trims a padded or dense matrix without materialising trailing null padding. */
function trimBounds(
	matrix: Matrix,
	rowMode: number,
	colMode: number,
): [number, number, number, number] {
	const blockRows = matrix.blockRows;
	const blockCols = matrix.blockCols;
	const padBlank = !matrix.isPadded || matrix.fill === null;
	if (!padBlank) {
		const data = matrix.data;
		const cols = data[0]?.length ?? 0;
		const [r0, r1] = trimmed(
			data.map((row) => row.every((v) => v === null)),
			rowMode,
		);
		const [c0, c1] = trimmed(
			Array.from({ length: cols }, (_, c) => data.every((row) => row[c] === null)),
			colMode,
		);
		return [r0, r1, c0, c1];
	}
	const rowBlank = Array.from({ length: blockRows }, (_, r) => {
		for (let c = 0; c < blockCols; c++) if (matrix.get(r, c) !== null) return false;
		return true;
	});
	const colBlank = Array.from({ length: blockCols }, (_, c) => {
		for (let r = 0; r < blockRows; r++) if (matrix.get(r, c) !== null) return false;
		return true;
	});
	let [r0, r1] = trimmed(rowBlank, rowMode);
	let [c0, c1] = trimmed(colBlank, colMode);
	if (rowMode === 0 || rowMode === 1) r1 = matrix.rows;
	if (colMode === 0 || colMode === 1) c1 = matrix.cols;
	return [r0, r1, c0, c1];
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
				const s = ctx.toScalar(v);
				const n = Math.trunc(s as number);
				if (typeof s !== 'number' || n < 0 || n > 3) fail(ERR.VALUE);
				return n;
			};
			const rowMode = mode(1);
			const colMode = mode(2);
			const source = args[0] ?? null;
			const matrix = ctx.toMatrix(source);
			const [r0, r1, c0, c1] = trimBounds(matrix, rowMode, colMode);
			if (r1 <= r0 || c1 <= c0) fail(ERR.CALC);
			if (source instanceof RefValue && source.areas.length === 1 && source.first) {
				const area = source.first;
				return new RefValue([
					{
						sheet: area.sheet,
						range: {
							start: { row: area.range.start.row + r0, col: area.range.start.col + c0 },
							end: { row: area.range.start.row + r1 - 1, col: area.range.start.col + c1 - 1 },
						},
					},
				]);
			}
			return Matrix.build(r1 - r0, c1 - c0, (r, c): Scalar => matrix.get(r0 + r, c0 + c));
		},
		['any'],
	),
];
