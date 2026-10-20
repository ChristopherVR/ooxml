// Which merges cover the cells of a sheet's rows, for rows visited top to bottom, in memory that
// follows the number of merges rather than merges times rows.
import type { CellRange } from '../address';

/** How the merges cover one cell. */
export interface MergeCover {
	/** Some merge covers the cell. */
	readonly merged: boolean;
	/** Some merge of more than one cell covers it. */
	readonly spanned: boolean;
}

const NONE: MergeCover = { merged: false, spanned: false };
const SINGLE: MergeCover = { merged: true, spanned: false };
const SPANNED: MergeCover = { merged: true, spanned: true };

const spans = (merge: CellRange): boolean =>
	merge.start.row !== merge.end.row || merge.start.col !== merge.end.col;

/** The index of the first of `merges` (ordered by first column) that starts right of `col`. */
function firstRightOf(merges: readonly CellRange[], col: number): number {
	let lo = 0;
	let hi = merges.length;
	while (lo < hi) {
		const mid = (lo + hi) >> 1;
		if ((merges[mid]?.start.col ?? 0) <= col) lo = mid + 1;
		else hi = mid;
	}
	return lo;
}

/**
 * Sweeps a sheet's rows top to bottom (`moveTo`) and tells how the merges cover a cell of the
 * current row (`at`). It holds only the merges covering the current row: a sheet's merges do not
 * overlap, so those are disjoint column spans, kept in column order and found by binary search.
 * Overlapping merges, which only a malformed file has, are detected and answered by a scan of the
 * current row's merges instead, which gives the same result.
 */
export class MergeSweep {
	/** The merges by first row; `next` is the first one the sweep has not reached. */
	private readonly byRow: CellRange[];
	private next = 0;
	private row = Number.NEGATIVE_INFINITY;
	/** The merges covering the current row, by first column. */
	private readonly active: CellRange[] = [];
	/** The last row of the earliest-ending active merge. */
	private firstEnd = Number.POSITIVE_INFINITY;
	private overlapping = false;

	constructor(merges: readonly CellRange[]) {
		this.byRow = [...merges].sort((a, b) => a.start.row - b.start.row);
	}

	/** How many merges cover the current row, which is everything the sweep holds. */
	get size(): number {
		return this.active.length;
	}

	/** Moves to `row`, which must not be above the previous one. */
	moveTo(row: number): void {
		if (row < this.row) throw new RangeError(`Rows must not go up: ${row} after ${this.row}`);
		this.row = row;
		if (row > this.firstEnd) this.dropEnded(row);
		for (; this.next < this.byRow.length; this.next++) {
			const merge = this.byRow[this.next];
			if (!merge || merge.start.row > row) break;
			if (merge.end.row >= row) this.add(merge);
		}
	}

	/** How the merges cover column `col` of the current row. */
	at(col: number): MergeCover {
		if (this.overlapping) {
			let merged = false;
			for (const merge of this.active) {
				if (col < merge.start.col || col > merge.end.col) continue;
				if (spans(merge)) return SPANNED;
				merged = true;
			}
			return merged ? SINGLE : NONE;
		}
		const merge = this.active[firstRightOf(this.active, col) - 1];
		if (!merge || col > merge.end.col) return NONE;
		return spans(merge) ? SPANNED : SINGLE;
	}

	private add(merge: CellRange): void {
		const at = firstRightOf(this.active, merge.start.col);
		// The active spans are disjoint and in column order, so a new one can only overlap its
		// neighbours.
		const before = this.active[at - 1];
		const after = this.active[at];
		if (
			(before && before.end.col >= merge.start.col) ||
			(after && merge.end.col >= after.start.col)
		)
			this.overlapping = true;
		this.active.splice(at, 0, merge);
		this.firstEnd = Math.min(this.firstEnd, merge.end.row);
	}

	private dropEnded(row: number): void {
		let kept = 0;
		let firstEnd = Number.POSITIVE_INFINITY;
		for (const merge of this.active) {
			if (merge.end.row < row) continue;
			this.active[kept++] = merge;
			firstEnd = Math.min(firstEnd, merge.end.row);
		}
		this.active.length = kept;
		this.firstEnd = firstEnd;
		if (!kept) this.overlapping = false;
	}
}
