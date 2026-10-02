import type { CallContext } from '../context.js';
import { ERR, fail, isError, Matrix, type Value } from '../values.js';
import { int, spec } from './helpers.js';
import type { FunctionSpec } from './types.js';

const C = 'Math & Trig';

/** A numeric matrix; text, blanks and logicals are #VALUE!. */
function numbers(ctx: CallContext, value: Value | undefined): number[][] {
	return ctx.toMatrix(value ?? null).data.map((row) =>
		row.map((v) => {
			if (isError(v)) fail(v);
			if (typeof v !== 'number') fail(ERR.VALUE);
			return v;
		}),
	);
}

function square(m: number[][]): number {
	const n = m.length;
	if (n === 0 || m.some((row) => row.length !== n)) fail(ERR.VALUE);
	return n;
}

/** Gaussian elimination with partial pivoting; returns the determinant and, optionally, the inverse. */
function eliminate(input: number[][], invert: boolean): { det: number; inverse?: number[][] } {
	const n = square(input);
	const a = input.map((row) => [...row]);
	const inv: number[][] = Array.from({ length: n }, (_, i) =>
		Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)),
	);
	let det = 1;
	for (let col = 0; col < n; col++) {
		let pivot = col;
		for (let r = col + 1; r < n; r++) {
			if (Math.abs(a[r]?.[col] ?? 0) > Math.abs(a[pivot]?.[col] ?? 0)) pivot = r;
		}
		const p = a[pivot]?.[col] ?? 0;
		if (Math.abs(p) < 1e-300) return { det: 0 };
		if (pivot !== col) {
			[a[pivot], a[col]] = [a[col] as number[], a[pivot] as number[]];
			[inv[pivot], inv[col]] = [inv[col] as number[], inv[pivot] as number[]];
			det = -det;
		}
		det *= p;
		const row = a[col] as number[];
		const irow = inv[col] as number[];
		for (let j = 0; j < n; j++) {
			row[j] = (row[j] as number) / p;
			irow[j] = (irow[j] as number) / p;
		}
		for (let r = 0; r < n; r++) {
			if (r === col) continue;
			const target = a[r] as number[];
			const itarget = inv[r] as number[];
			const factor = target[col] as number;
			if (factor === 0) continue;
			for (let j = 0; j < n; j++) {
				target[j] = (target[j] as number) - factor * (row[j] as number);
				itarget[j] = (itarget[j] as number) - factor * (irow[j] as number);
			}
		}
	}
	return invert ? { det, inverse: inv } : { det };
}

export const MATRIX_FUNCTIONS: FunctionSpec[] = [
	spec(
		'MMULT',
		C,
		'MMULT(array1, array2)',
		'The matrix product of two arrays.',
		2,
		2,
		(args, ctx) => {
			const a = numbers(ctx, args[0]);
			const b = numbers(ctx, args[1]);
			const inner = a[0]?.length ?? 0;
			if (inner !== b.length) fail(ERR.VALUE);
			const cols = b[0]?.length ?? 0;
			return Matrix.build(a.length, cols, (r, c) => {
				let total = 0;
				for (let k = 0; k < inner; k++) total += (a[r]?.[k] ?? 0) * (b[k]?.[c] ?? 0);
				return total;
			});
		},
		['any'],
	),
	spec(
		'MDETERM',
		C,
		'MDETERM(array)',
		'The determinant of a square matrix.',
		1,
		1,
		(args, ctx) => eliminate(numbers(ctx, args[0]), false).det,
		['any'],
	),
	spec(
		'MINVERSE',
		C,
		'MINVERSE(array)',
		'The inverse of a square matrix.',
		1,
		1,
		(args, ctx) => {
			const { det, inverse } = eliminate(numbers(ctx, args[0]), true);
			if (det === 0 || !inverse) fail(ERR.NUM);
			return new Matrix(inverse);
		},
		['any'],
	),
	spec('MUNIT', C, 'MUNIT(dimension)', 'The identity matrix.', 1, 1, (args) => {
		const n = int(args[0]);
		if (n < 1) fail(ERR.VALUE);
		return Matrix.build(n, n, (r, c) => (r === c ? 1 : 0));
	}),
];
