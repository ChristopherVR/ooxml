// Shared plumbing for SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS and MINIFS.
import type { CallContext } from '../context.js';
import { ERR, fail, Matrix, RefValue, type Scalar, type Value } from '../values.js';
import { makeCriteria, scalar } from './helpers.js';

const CLIP = 100_000;

/**
 * Reads several range arguments with one shape. With `sameShape`, every argument must have the
 * first one's size (the IFS functions); otherwise later ranges are resized from their top-left
 * cell (SUMIF's sum_range). Whole-column references are clipped to the used rows.
 */
export function shapedMatrices(ctx: CallContext, values: Value[], sameShape: boolean): Matrix[] {
	const first = values[0];
	let rows: number;
	let cols: number;
	if (first instanceof RefValue) {
		const area = first.areas[0];
		if (!area || first.areas.length !== 1) fail(ERR.VALUE);
		rows = area.range.end.row - area.range.start.row + 1;
		cols = area.range.end.col - area.range.start.col + 1;
	} else if (first instanceof Matrix) {
		rows = first.rows;
		cols = first.cols;
	} else {
		rows = 1;
		cols = 1;
	}
	if (sameShape) {
		for (const value of values) {
			const shape = shapeOf(value);
			if (shape.rows !== rows || shape.cols !== cols) fail(ERR.VALUE);
		}
	}
	if (rows * cols > CLIP) {
		let maxRows = 1;
		let maxCols = 1;
		for (const value of values) {
			if (!(value instanceof RefValue)) continue;
			const area = value.areas[0];
			if (!area) continue;
			const bounds = ctx.frame.host.bounds(area.sheet);
			maxRows = Math.max(maxRows, bounds.rows - area.range.start.row);
			maxCols = Math.max(maxCols, bounds.cols - area.range.start.col);
		}
		rows = Math.min(rows, maxRows);
		cols = Math.min(cols, maxCols);
	}
	return values.map((value) => {
		if (value instanceof RefValue) {
			const area = value.areas[0];
			if (!area || value.areas.length !== 1) fail(ERR.VALUE);
			const { row, col } = area.range.start;
			return Matrix.build(rows, cols, (r, c) => ctx.readCell(area.sheet, row + r, col + c));
		}
		if (value instanceof Matrix) return value;
		return new Matrix([[scalar(value)]]);
	});
}

function shapeOf(value: Value): { rows: number; cols: number } {
	if (value instanceof RefValue) {
		const area = value.areas[0];
		if (!area) return { rows: 0, cols: 0 };
		return {
			rows: area.range.end.row - area.range.start.row + 1,
			cols: area.range.end.col - area.range.start.col + 1,
		};
	}
	if (value instanceof Matrix) return { rows: value.rows, cols: value.cols };
	return { rows: 1, cols: 1 };
}

/**
 * For `(range, criteria)` pairs, the values of `target` (or of the first range) at positions
 * where every criterion holds.
 */
export function matchingValues(
	ctx: CallContext,
	pairs: { range: Value; criteria: Value }[],
	target: Value | undefined,
	sameShape: boolean,
): Scalar[] {
	const ranges = pairs.map((p) => p.range);
	const matrices = shapedMatrices(
		ctx,
		target === undefined ? ranges : [...ranges, target],
		sameShape,
	);
	const tests = pairs.map((p) => makeCriteria(scalar(p.criteria)));
	const values = matrices[matrices.length - 1] as Matrix;
	const out: Scalar[] = [];
	const first = matrices[0] as Matrix;
	for (let r = 0; r < first.rows; r++) {
		for (let c = 0; c < first.cols; c++) {
			let ok = true;
			for (let i = 0; i < tests.length && ok; i++) {
				ok = (tests[i] as (v: Scalar) => boolean)((matrices[i] as Matrix).get(r, c));
			}
			if (ok) out.push(values.get(r, c));
		}
	}
	return out;
}

/** Splits trailing `(range, criteria)` argument pairs. */
export function criteriaPairs(args: Value[], from: number): { range: Value; criteria: Value }[] {
	if ((args.length - from) % 2 !== 0) fail(ERR.VALUE);
	const pairs: { range: Value; criteria: Value }[] = [];
	for (let i = from; i < args.length; i += 2) {
		pairs.push({ range: args[i] ?? null, criteria: args[i + 1] ?? null });
	}
	return pairs;
}
