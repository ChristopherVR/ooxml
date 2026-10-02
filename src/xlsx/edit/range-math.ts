import { MAX_COL, MAX_ROW, type CellAddress, type CellRange, rangesIntersect } from '../address.js';

export type Axis = 'row' | 'col';

/** An axis shift: `count` positive inserts before index `at`, negative deletes `-count` from `at`. */
export interface AxisShift {
	axis: Axis;
	at: number;
	count: number;
}

const axisMax = (axis: Axis): number => (axis === 'row' ? MAX_ROW : MAX_COL);

/**
 * Shifts a one-dimensional inclusive span. Inserting at the start moves the span; inserting inside
 * it widens it. Deleting trims it; a span deleted entirely returns undefined.
 */
export function shiftSpan(
	start: number,
	end: number,
	shift: AxisShift,
): { start: number; end: number } | undefined {
	const { at, count } = shift;
	const max = axisMax(shift.axis);
	if (count > 0) {
		if (start >= at) {
			if (start + count > max) return undefined;
			return { start: start + count, end: Math.min(max, end + count) };
		}
		if (end >= at) return { start, end: Math.min(max, end + count) };
		return { start, end };
	}
	const n = -count;
	const last = at + n - 1;
	if (start >= at && end <= last) return undefined;
	const s = start < at ? start : start > last ? start - n : at;
	// A span reaching the grid edge (a whole row or column reference) keeps reaching it.
	const e = end >= max ? max : end < at ? end : end > last ? end - n : at - 1;
	return s > e ? undefined : { start: s, end: e };
}

/** Shifts one index; undefined when it was deleted (or pushed off the grid). */
export function shiftIndex(index: number, shift: AxisShift): number | undefined {
	const { at, count } = shift;
	if (index < at) return index;
	if (count > 0) return index + count > axisMax(shift.axis) ? undefined : index + count;
	return index < at - count ? undefined : index + count;
}

/** Shifts an index, clamping deleted positions to the start of the deleted band (drawing anchors). */
export function shiftIndexClamped(index: number, shift: AxisShift): number {
	const moved = shiftIndex(index, shift);
	if (moved !== undefined) return moved;
	return shift.count > 0 ? axisMax(shift.axis) : shift.at;
}

export function shiftRange(range: CellRange, shift: AxisShift): CellRange | undefined {
	if (shift.axis === 'row') {
		const span = shiftSpan(range.start.row, range.end.row, shift);
		return (
			span && {
				start: { row: span.start, col: range.start.col },
				end: { row: span.end, col: range.end.col },
			}
		);
	}
	const span = shiftSpan(range.start.col, range.end.col, shift);
	return (
		span && {
			start: { row: range.start.row, col: span.start },
			end: { row: range.end.row, col: span.end },
		}
	);
}

export function shiftAddress(address: CellAddress, shift: AxisShift): CellAddress | undefined {
	const key = shift.axis === 'row' ? 'row' : 'col';
	const moved = shiftIndex(address[key], shift);
	return moved === undefined ? undefined : { ...address, [key]: moved };
}

/**
 * Shifts a range inside a band (only the cells between `lo` and `hi` on the other axis move, as an
 * insert-cells or delete-cells shift does). Ranges straddling the band edge stay where they are.
 */
export function shiftRangeInBand(
	range: CellRange,
	shift: AxisShift,
	lo: number,
	hi: number,
): CellRange | undefined {
	const other = shift.axis === 'row' ? 'col' : 'row';
	if (range.end[other] < lo || range.start[other] > hi) return range;
	if (range.start[other] >= lo && range.end[other] <= hi) return shiftRange(range, shift);
	return range;
}

/** `a` minus `b`: up to four rectangles covering what of `a` lies outside `b`. */
export function subtractRange(a: CellRange, b: CellRange): CellRange[] {
	if (!rangesIntersect(a, b)) return [a];
	const out: CellRange[] = [];
	const top = Math.max(a.start.row, b.start.row);
	const bottom = Math.min(a.end.row, b.end.row);
	if (a.start.row < b.start.row)
		out.push({ start: a.start, end: { row: b.start.row - 1, col: a.end.col } });
	if (a.end.row > b.end.row)
		out.push({ start: { row: b.end.row + 1, col: a.start.col }, end: a.end });
	if (a.start.col < b.start.col)
		out.push({ start: { row: top, col: a.start.col }, end: { row: bottom, col: b.start.col - 1 } });
	if (a.end.col > b.end.col)
		out.push({ start: { row: top, col: b.end.col + 1 }, end: { row: bottom, col: a.end.col } });
	return out;
}

export const sameRange = (a: CellRange, b: CellRange): boolean =>
	a.start.row === b.start.row &&
	a.start.col === b.start.col &&
	a.end.row === b.end.row &&
	a.end.col === b.end.col;

export const rangeWithin = (inner: CellRange, outer: CellRange): boolean =>
	inner.start.row >= outer.start.row &&
	inner.end.row <= outer.end.row &&
	inner.start.col >= outer.start.col &&
	inner.end.col <= outer.end.col;

export const cellRange = (row: number, col: number, row2 = row, col2 = col): CellRange => ({
	start: { row, col },
	end: { row: row2, col: col2 },
});
