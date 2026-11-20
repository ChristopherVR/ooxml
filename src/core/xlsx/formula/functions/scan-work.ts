// What a scan of a range costs of a formula's work budget, for functions an array lifts over its
// elements: each element repeats the scan, which is how one formula multiplies its work. A unit is
// about a nanosecond (see `MAX_FORMULA_WORK`); the prices are this engine's measured costs.
import type { CallContext } from '../context';
import { Matrix, RefValue, type Value } from '../values';

/** Testing one cell against a criterion (COUNTIF, SUMIFS and the other criteria functions). */
export const CELL_TEST_WORK = 27;
/** Comparing one cell with a lookup value (XLOOKUP, XMATCH, MATCH and the other lookups). */
export const LOOKUP_TEST_WORK = 20;

/** Cells a range argument is scanned over: a reference only as far as its sheet's used part. */
export function scannedCells(ctx: CallContext, range: Value): number {
	if (range instanceof Matrix) return range.blockRows * range.blockCols;
	if (!(range instanceof RefValue)) return 1;
	let cells = 0;
	for (const { sheet, range: r } of range.areas) {
		const used = ctx.frame.host.bounds(sheet);
		const rows = Math.min(r.end.row, used.rows - 1) - r.start.row + 1;
		const cols = Math.min(r.end.col, used.cols - 1) - r.start.col + 1;
		if (rows > 0 && cols > 0) cells += rows * cols;
	}
	return cells;
}

/** The work of a linear scan of `range`, `price` per cell. */
export const linearScanWork = (ctx: CallContext, range: Value, price = CELL_TEST_WORK): number =>
	scannedCells(ctx, range) * price;

/** The work of a binary search of `range`. */
export const binarySearchWork = (ctx: CallContext, range: Value): number =>
	Math.ceil(Math.log2(scannedCells(ctx, range) + 1)) * LOOKUP_TEST_WORK;
