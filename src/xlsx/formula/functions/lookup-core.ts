// Vector access and search shared by MATCH, XMATCH, VLOOKUP, HLOOKUP, LOOKUP and XLOOKUP.
import { compareScalars } from '../coerce.js';
import type { CallContext } from '../context.js';
import { ERR, fail, isError, Matrix, RefValue, type Scalar, type Value } from '../values.js';
import { hasWildcards, wildcardRegex } from './helpers.js';

/** A lazily read row or column of values. */
export interface Vector {
	length: number;
	get(index: number): Scalar;
}

/** The dimensions of a reference or array argument. */
export function shape(value: Value): { rows: number; cols: number } {
	if (value instanceof RefValue) {
		const area = value.areas[0];
		if (!area || value.areas.length !== 1) fail(ERR.VALUE);
		return {
			rows: area.range.end.row - area.range.start.row + 1,
			cols: area.range.end.col - area.range.start.col + 1,
		};
	}
	if (value instanceof Matrix) return { rows: value.rows, cols: value.cols };
	return { rows: 1, cols: 1 };
}

/** Row `index` (`axis` 'row') or column `index` of a value, read on demand and clipped to used cells. */
export function line(ctx: CallContext, value: Value, axis: 'row' | 'col', index: number): Vector {
	if (value instanceof RefValue) {
		const area = value.areas[0];
		if (!area || value.areas.length !== 1) fail(ERR.VALUE);
		const { start, end } = area.range;
		const bounds = ctx.frame.host.bounds(area.sheet);
		if (axis === 'col') {
			const col = start.col + index;
			const length = Math.max(0, Math.min(end.row, bounds.rows - 1) - start.row + 1);
			return { length, get: (i) => ctx.readCell(area.sheet, start.row + i, col) };
		}
		const row = start.row + index;
		const length = Math.max(0, Math.min(end.col, bounds.cols - 1) - start.col + 1);
		return { length, get: (i) => ctx.readCell(area.sheet, row, start.col + i) };
	}
	const m = value instanceof Matrix ? value : new Matrix([[value as Scalar]]);
	if (axis === 'col') return { length: m.rows, get: (i) => m.get(i, index) };
	return { length: m.cols, get: (i) => m.get(index, i) };
}

/** A one-dimensional argument as a vector (a single row or column); 2D is #N/A. */
export function vectorOf(ctx: CallContext, value: Value): { vector: Vector; axis: 'row' | 'col' } {
	const { rows, cols } = shape(value);
	if (rows !== 1 && cols !== 1) fail(ERR.NA);
	if (cols === 1 && rows >= 1) return { vector: line(ctx, value, 'col', 0), axis: 'col' };
	return { vector: line(ctx, value, 'row', 0), axis: 'row' };
}

const sameKind = (a: Scalar, b: Scalar): boolean =>
	(typeof a === 'number' && typeof b === 'number') ||
	(typeof a === 'string' && typeof b === 'string') ||
	(typeof a === 'boolean' && typeof b === 'boolean');

/** An exact-match predicate (case-insensitive text, optional wildcards). */
export function exactMatcher(lookup: Scalar, wildcards: boolean): (v: Scalar) => boolean {
	if (typeof lookup === 'string' && wildcards && hasWildcards(lookup)) {
		const regex = wildcardRegex(lookup);
		return (v) => typeof v === 'string' && regex.test(v);
	}
	if (lookup === null) return (v) => v === null || v === '';
	return (v) =>
		!isError(v) && v !== null && sameKind(v, lookup) && compareScalars(v, lookup as never) === 0;
}

export function findExact(
	vector: Vector,
	lookup: Scalar,
	wildcards: boolean,
	reverse = false,
): number {
	const test = exactMatcher(lookup, wildcards);
	if (reverse) {
		for (let i = vector.length - 1; i >= 0; i--) if (test(vector.get(i))) return i;
		return -1;
	}
	for (let i = 0; i < vector.length; i++) if (test(vector.get(i))) return i;
	return -1;
}

/**
 * Binary search like Excel's approximate lookups: on ascending data the last position whose value
 * is <= `lookup` (`descending`: the last position whose value is >= `lookup`).
 */
export function findSorted(vector: Vector, lookup: Scalar, descending = false): number {
	if (isError(lookup)) fail(lookup);
	let lo = 0;
	let hi = vector.length - 1;
	let best = -1;
	while (lo <= hi) {
		const mid = (lo + hi) >> 1;
		let v = vector.get(mid);
		// Blank and error cells are stepped over toward the low end.
		let probe = mid;
		while ((v === null || isError(v)) && probe > lo) {
			probe--;
			v = vector.get(probe);
		}
		if (v === null || isError(v) || !sameKind(v, lookup)) {
			if (v !== null && !isError(v) && compareScalars(v, lookup as never) > 0 !== descending) {
				hi = probe - 1;
			} else {
				lo = mid + 1;
			}
			continue;
		}
		const c = compareScalars(v, lookup as never);
		if (descending ? c >= 0 : c <= 0) {
			best = probe;
			lo = mid + 1;
		} else {
			hi = probe - 1;
		}
	}
	return best;
}

/** XLOOKUP / XMATCH next-smaller (-1) or next-larger (1) match by a linear scan. */
export function findNearest(
	vector: Vector,
	lookup: Scalar,
	mode: -1 | 1,
	reverse: boolean,
): number {
	let best = -1;
	let bestValue: Scalar = null;
	const n = vector.length;
	for (let k = 0; k < n; k++) {
		const i = reverse ? n - 1 - k : k;
		const v = vector.get(i);
		if (v === null || isError(v) || !sameKind(v, lookup)) continue;
		const c = compareScalars(v, lookup as never);
		if (c === 0) return i;
		if (mode === -1 ? c < 0 : c > 0) {
			if (
				best < 0 ||
				(mode === -1
					? compareScalars(v, bestValue as never) > 0
					: compareScalars(v, bestValue as never) < 0)
			) {
				best = i;
				bestValue = v;
			}
		}
	}
	return best;
}

/** The XLOOKUP / XMATCH search. Returns the 0-based index or -1. */
export function xsearch(
	vector: Vector,
	lookup: Scalar,
	matchMode: number,
	searchMode: number,
): number {
	if (![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) fail(ERR.VALUE);
	if (searchMode === 2 || searchMode === -2) {
		const descending = searchMode === -2;
		const at = findSorted(vector, lookup, descending);
		if (matchMode === 0 || matchMode === 2) {
			return at >= 0 && exactMatcher(lookup, false)(vector.get(at)) ? at : -1;
		}
		if (at >= 0 && exactMatcher(lookup, false)(vector.get(at))) return at;
		if (matchMode === -1) return descending ? (at + 1 < vector.length ? at + 1 : -1) : at;
		return descending ? at : at + 1 < vector.length ? at + 1 : -1;
	}
	const reverse = searchMode === -1;
	if (matchMode === 0 || matchMode === 2)
		return findExact(vector, lookup, matchMode === 2, reverse);
	return findNearest(vector, lookup, matchMode as -1 | 1, reverse);
}
