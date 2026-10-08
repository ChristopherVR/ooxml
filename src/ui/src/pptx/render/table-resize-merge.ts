/**
 * table-resize-merge.ts - which stretches of a table's internal boundaries are
 * real cell edges once merged cells are taken into account.
 *
 * A row boundary inside a vertically merged cell (or a column boundary inside
 * a horizontally merged one) is not a visible edge, and PowerPoint offers no
 * resize there. Drawing a full-width row handle across a `rowspan` cell put a
 * resize target over the middle of that cell, so a click on the merged cell
 * started a row drag instead of selecting it. Every binding's resize overlay
 * uses these helpers to draw handles only along real edges and to ignore a
 * press that lands on a boundary inside a merge.
 *
 * Pure and framework-free. Cells are indexed by grid column, the way the
 * parser stores them (continuation cells are present with `hMerge` / `vMerge`).
 */
import type { PptxTableRow } from 'ooxml-core/pptx';

/** An inclusive run of grid indexes along which a boundary is a real edge. */
export interface TableBoundaryRun {
	start: number;
	end: number;
}

/** A row-boundary handle segment, as percentages of the table width. */
export interface TableRowHandleSegment {
	leftPct: number;
	widthPct: number;
}

/**
 * Where merged cells cross the internal boundaries.
 * `rows[b][c]` is true when a merge spans row boundary `b` (between rows `b`
 * and `b + 1`) at grid column `c`; `columns[b][r]` is true when a merge spans
 * column boundary `b` (between columns `b` and `b + 1`) at row `r`.
 */
export interface TableMergeCrossings {
	rows: boolean[][];
	columns: boolean[][];
}

/** Map every internal boundary to the grid lines a merged cell spans it at. */
export function computeTableMergeCrossings(
	rows: readonly PptxTableRow[],
	columnCount: number,
): TableMergeCrossings {
	const rowCount = rows.length;
	const rowCrossings = Array.from({ length: Math.max(0, rowCount - 1) }, () =>
		Array.from({ length: columnCount }, () => false),
	);
	const columnCrossings = Array.from({ length: Math.max(0, columnCount - 1) }, () =>
		Array.from({ length: rowCount }, () => false),
	);
	rows.forEach((row, r) => {
		row.cells.forEach((cell, c) => {
			if (cell.hMerge || cell.vMerge) {
				return;
			}
			const rowEnd = Math.min(rowCount, r + Math.max(1, cell.rowSpan ?? 1)) - 1;
			const colEnd = Math.min(columnCount, c + Math.max(1, cell.gridSpan ?? 1)) - 1;
			for (let rr = r; rr <= rowEnd; rr++) {
				for (let cc = c; cc <= colEnd; cc++) {
					if (rr < rowEnd) {
						rowCrossings[rr]![cc] = true;
					}
					if (cc < colEnd) {
						columnCrossings[cc]![rr] = true;
					}
				}
			}
		});
	});
	return { rows: rowCrossings, columns: columnCrossings };
}

/** The maximal runs of indexes a boundary is NOT crossed at (its real edges). */
export function openBoundaryRuns(crossed: readonly boolean[]): TableBoundaryRun[] {
	const runs: TableBoundaryRun[] = [];
	let start = -1;
	crossed.forEach((isCrossed, i) => {
		if (!isCrossed && start < 0) {
			start = i;
		} else if (isCrossed && start >= 0) {
			runs.push({ start, end: i - 1 });
			start = -1;
		}
	});
	if (start >= 0) {
		runs.push({ start, end: crossed.length - 1 });
	}
	return runs;
}

/** Row-boundary handle segments for one boundary, positioned by column widths. */
export function rowHandleSegments(
	crossed: readonly boolean[],
	columnWidths: readonly number[],
): TableRowHandleSegment[] {
	const total = columnWidths.reduce((sum, width) => sum + width, 0) || 1;
	const offset = (index: number): number =>
		(columnWidths.slice(0, index).reduce((sum, width) => sum + width, 0) / total) * 100;
	// A column `crossed` does not cover (no merge information) counts as an edge.
	const edges = columnWidths.map((_, i) => crossed[i] === true);
	return openBoundaryRuns(edges).map(({ start, end }) => {
		const leftPct = offset(start);
		return { leftPct, widthPct: offset(end + 1) - leftPct };
	});
}

/**
 * Column-boundary handle segments for one boundary, in pixels from the table
 * top, given the measured internal row boundaries and the table height.
 */
export function columnHandleSegments(
	crossed: readonly boolean[],
	rowBounds: readonly number[],
	tableHeight: number,
): { top: number; height: number }[] {
	// A row `crossed` does not cover (no merge information) counts as an edge.
	const edges = Array.from({ length: rowBounds.length + 1 }, (_, i) => crossed[i] === true);
	return openBoundaryRuns(edges).map(({ start, end }) => {
		const top = start === 0 ? 0 : (rowBounds[start - 1] ?? 0);
		const bottom = rowBounds[end] ?? tableHeight;
		return { top, height: Math.max(0, bottom - top) };
	});
}

/** The grid column under a horizontal position given as a fraction (0-1) of the table width. */
export function columnIndexAtFraction(columnWidths: readonly number[], fraction: number): number {
	const total = columnWidths.reduce((sum, width) => sum + width, 0) || 1;
	let cumulative = 0;
	for (let i = 0; i < columnWidths.length; i++) {
		cumulative += columnWidths[i]! / total;
		if (fraction < cumulative) {
			return i;
		}
	}
	return Math.max(0, columnWidths.length - 1);
}

/** The row under a vertical offset, given the internal row boundaries in the same units. */
export function rowIndexAtOffset(rowBounds: readonly number[], offset: number): number {
	const index = rowBounds.findIndex((bound) => offset < bound);
	return index < 0 ? rowBounds.length : index;
}

/** Whether a press at `xFraction` (0-1 of the width) on row boundary `boundary` falls inside a merge. */
export function isRowBoundaryMergedAt(
	crossings: TableMergeCrossings,
	boundary: number,
	columnWidths: readonly number[],
	xFraction: number,
): boolean {
	return crossings.rows[boundary]?.[columnIndexAtFraction(columnWidths, xFraction)] === true;
}

/** Whether a press at `offset` (same units as `rowBounds`) on column boundary `boundary` falls inside a merge. */
export function isColumnBoundaryMergedAt(
	crossings: TableMergeCrossings,
	boundary: number,
	rowBounds: readonly number[],
	offset: number,
): boolean {
	return crossings.columns[boundary]?.[rowIndexAtOffset(rowBounds, offset)] === true;
}
