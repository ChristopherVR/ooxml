// Shared plumbing for SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS and MINIFS.
import type { CellRange } from '../../address.js';
import type { CallContext } from '../context.js';
import { pick } from '../operators.js';
import {
	type Area,
	ERR,
	ErrorSignal,
	fail,
	LambdaValue,
	Matrix,
	RefValue,
	type Scalar,
	type Value,
} from '../values.js';
import { equalityKey, positionsOf } from './criteria-index.js';
import { makeCriteria, scalar } from './helpers.js';

const CLIP = 100_000;

export interface CriteriaPair {
	range: Value;
	criteria: Value;
}

/** Cells matched by a set of criteria. */
export interface Matches {
	/** Values of the target (or the first range) where every criterion holds, in the stored part. */
	values: Scalar[];
	/** Matching cells past the sheet's used part: all blank, so they are counted without reading. */
	blankTail: number;
}

function areaOf(value: RefValue): Area {
	const area = value.areas[0];
	if (!area || value.areas.length !== 1) fail(ERR.VALUE);
	return area;
}

function shapeOf(value: Value): { rows: number; cols: number } {
	if (value instanceof RefValue) {
		const { start, end } = areaOf(value).range;
		return { rows: end.row - start.row + 1, cols: end.col - start.col + 1 };
	}
	if (value instanceof Matrix) return { rows: value.rows, cols: value.cols };
	return { rows: 1, cols: 1 };
}

/** The values of `rows` x `cols` cells from an area's top-left cell (cached during a recalc). */
function readBlock(ctx: CallContext, area: Area, rows: number, cols: number): Matrix {
	const { row, col } = area.range.start;
	const range: CellRange = {
		start: { row, col },
		end: { row: row + rows - 1, col: col + cols - 1 },
	};
	const host = ctx.frame.host;
	if (host.readBlock) return host.readBlock(area.sheet, range);
	return Matrix.build(rows, cols, (r, c) => ctx.readCell(area.sheet, row + r, col + c));
}

/**
 * Reads several range arguments with one shape. With `sameShape`, every argument must have the
 * first one's size (the IFS functions); otherwise later ranges are resized from their top-left
 * cell (SUMIF's sum_range). Whole-column references read only the used rows: `total` is the full
 * cell count and `block` the part read.
 */
function shapedMatrices(
	ctx: CallContext,
	values: Value[],
	sameShape: boolean,
): { matrices: Matrix[]; total: number; block: number } {
	let { rows, cols } = shapeOf(values[0] ?? null);
	if (sameShape) {
		for (const value of values) {
			const shape = shapeOf(value);
			if (shape.rows !== rows || shape.cols !== cols) fail(ERR.VALUE);
		}
	}
	const total = rows * cols;
	if (total > CLIP && values.every((v) => v instanceof RefValue)) {
		let maxRows = 1;
		let maxCols = 1;
		for (const value of values) {
			const area = areaOf(value as RefValue);
			const bounds = ctx.frame.host.bounds(area.sheet);
			maxRows = Math.max(maxRows, bounds.rows - area.range.start.row);
			maxCols = Math.max(maxCols, bounds.cols - area.range.start.col);
		}
		rows = Math.min(rows, maxRows);
		cols = Math.min(cols, maxCols);
	}
	const matrices = values.map((value) => {
		if (value instanceof RefValue) return readBlock(ctx, areaOf(value), rows, cols);
		if (value instanceof Matrix) return value;
		return new Matrix([[scalar(value)]]);
	});
	return { matrices, total, block: rows * cols };
}

/**
 * For `(range, criteria)` pairs with single criteria, the values of `target` (or of the first
 * range) at positions where every criterion holds, plus how many matching cells lie past the
 * used part of whole-column ranges.
 */
export function matchingCells(
	ctx: CallContext,
	pairs: CriteriaPair[],
	target: Value | undefined,
	sameShape: boolean,
): Matches {
	const ranges = pairs.map((p) => p.range);
	const { matrices, total, block } = shapedMatrices(
		ctx,
		target === undefined ? ranges : [...ranges, target],
		sameShape,
	);
	const tests = pairs.map((p) => makeCriteria(scalar(p.criteria)));
	const values = matrices[matrices.length - 1] as Matrix;
	const out: Scalar[] = [];
	const first = matrices[0] as Matrix;
	const holds = (r: number, c: number): boolean => {
		for (let i = 0; i < tests.length; i++) {
			if (!(tests[i] as (v: Scalar) => boolean)((matrices[i] as Matrix).get(r, c))) return false;
		}
		return true;
	};
	const candidates = indexedCandidates(pairs, matrices);
	if (candidates) {
		for (const at of candidates) {
			const r = Math.floor(at / first.cols);
			const c = at - r * first.cols;
			if (holds(r, c)) out.push(values.get(r, c));
		}
	} else {
		for (let r = 0; r < first.rows; r++) {
			for (let c = 0; c < first.cols; c++) if (holds(r, c)) out.push(values.get(r, c));
		}
	}
	const blankTail = total > block && tests.every((test) => test(null)) ? total - block : 0;
	return { values: out, blankTail };
}

/** Positions an equality criterion allows, from an index of a range block read repeatedly. */
function indexedCandidates(pairs: CriteriaPair[], matrices: Matrix[]): number[] | undefined {
	const first = matrices[0] as Matrix;
	for (let i = 0; i < pairs.length; i++) {
		const m = matrices[i] as Matrix;
		if (!(pairs[i]?.range instanceof RefValue) || m.cols !== first.cols) continue;
		const key = equalityKey(scalar(pairs[i]?.criteria));
		if (key === undefined) continue;
		const found = positionsOf(m, key);
		if (found) return found;
	}
	return undefined;
}

/** The matching values (see `matchingCells`); blanks past the used part add nothing to them. */
export const matchingValues = (
	ctx: CallContext,
	pairs: CriteriaPair[],
	target: Value | undefined,
	sameShape: boolean,
): Scalar[] => matchingCells(ctx, pairs, target, sameShape).values;

/** Splits trailing `(range, criteria)` argument pairs. */
export function criteriaPairs(args: Value[], from: number): CriteriaPair[] {
	if ((args.length - from) % 2 !== 0) fail(ERR.VALUE);
	const pairs: CriteriaPair[] = [];
	for (let i = from; i < args.length; i += 2) {
		pairs.push({ range: args[i] ?? null, criteria: args[i + 1] ?? null });
	}
	return pairs;
}

/** A criteria argument as one value, or as an array when it holds several. */
function criteriaValue(ctx: CallContext, value: Value): Scalar | Matrix {
	if (value instanceof LambdaValue) fail(ERR.CALC);
	let v: Value = value;
	if (v instanceof RefValue) {
		if (v.isCell()) return ctx.toScalar(v);
		v = ctx.toMatrix(v);
	}
	if (v instanceof Matrix) return v.rows === 1 && v.cols === 1 ? v.get(0, 0) : v;
	return v as Scalar;
}

/**
 * Runs an IFS function. An array (or multi-cell range) of criteria lifts the call like Excel:
 * the result is an array with one result per criterion, so SUM(COUNTIFS(B:B,{"a","b"})) adds
 * the counts of both.
 */
export function liftCriteria(
	ctx: CallContext,
	pairs: CriteriaPair[],
	fn: (pairs: CriteriaPair[]) => Value,
): Value {
	const resolved = pairs.map((p) => ({ range: p.range, criteria: criteriaValue(ctx, p.criteria) }));
	const arrays = resolved.map((p) => p.criteria).filter((c): c is Matrix => c instanceof Matrix);
	if (arrays.length === 0) return fn(resolved);
	let rows = 1;
	let cols = 1;
	for (const m of arrays) {
		rows = Math.max(rows, m.rows);
		cols = Math.max(cols, m.cols);
	}
	return Matrix.build(rows, cols, (r, c) => {
		const one = resolved.map((p) => ({
			range: p.range,
			criteria: p.criteria instanceof Matrix ? pick(p.criteria, r, c) : p.criteria,
		}));
		try {
			const value = fn(one);
			return value instanceof Matrix ? value.get(0, 0) : scalar(value);
		} catch (e) {
			if (e instanceof ErrorSignal) return e.value;
			throw e;
		}
	});
}
