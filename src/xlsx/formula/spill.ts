// Dynamic-array spilling: results written into neighbouring cells, marked so they are not
// mistaken for user data (and not persisted as constants on save).
import { type CellAddress, type CellRange, MAX_COL, MAX_ROW, rangesIntersect } from '../address.js';
import { deleteCell, getCell, putCell } from '../cells.js';
import type { Cell, CellValue, Worksheet } from '../model.js';
import { pick } from './operators.js';
import type { Matrix } from './values.js';

/** A cell whose value was written by a spilling formula anchored at `spillAnchor`. */
export interface SpilledCell extends Cell {
	/** The formula cell that owns this value (same sheet). */
	spillAnchor: CellAddress;
}

/**
 * Whether a cell's value is a dynamic-array spill result rather than user data. Writers should
 * not persist such values as constants (the anchor's formula recreates them); a spilled cell may
 * still carry its own style.
 */
export const isSpilledCell = (cell: Cell | undefined): cell is SpilledCell =>
	cell !== undefined && 'spillAnchor' in cell;

/** The anchor of a spilled cell, if it is one. */
export const spillAnchorOf = (cell: Cell | undefined): CellAddress | undefined =>
	isSpilledCell(cell) ? { ...cell.spillAnchor } : undefined;

/** Removes the spill marker and value from a cell (deleting it when nothing else remains). */
export function releaseSpilledCell(sheet: Worksheet, row: number, col: number): void {
	const cell = getCell(sheet, row, col);
	if (!isSpilledCell(cell)) return;
	delete (cell as Partial<SpilledCell>).spillAnchor;
	cell.value = null;
	if (!cell.styleId && !cell.formula && !cell.richText && !cell.arrayRange)
		deleteCell(sheet, row, col);
}

/** Clears the cells of a spill footprint that still belong to the anchor at (row, col). */
export function clearFootprint(
	sheet: Worksheet,
	footprint: CellRange,
	row: number,
	col: number,
): void {
	for (let r = footprint.start.row; r <= footprint.end.row; r++) {
		const cells = sheet.rows.get(r);
		if (!cells) continue;
		for (let c = footprint.start.col; c <= footprint.end.col; c++) {
			if (r === row && c === col) continue;
			const cell = cells.get(c);
			if (isSpilledCell(cell) && cell.spillAnchor.row === row && cell.spillAnchor.col === col) {
				releaseSpilledCell(sheet, r, c);
			}
		}
	}
}

/** The footprint an array result anchored at (row, col) needs, or undefined beyond the grid. */
export function footprintFor(row: number, col: number, m: Matrix): CellRange | undefined {
	const end = { row: row + m.rows - 1, col: col + m.cols - 1 };
	if (end.row > MAX_ROW || end.col > MAX_COL) return undefined;
	return { start: { row, col }, end };
}

/**
 * Whether a footprint is free: every other cell is empty, a spill value of this anchor, or a
 * stale spill value whose anchor no longer claims it (`claims`). Merged cells block.
 */
export function footprintFree(
	sheet: Worksheet,
	footprint: CellRange,
	row: number,
	col: number,
	claims: (anchor: CellAddress, r: number, c: number) => boolean,
): boolean {
	if (sheet.merges.some((m) => rangesIntersect(m, footprint))) return false;
	for (let r = footprint.start.row; r <= footprint.end.row; r++) {
		const cells = sheet.rows.get(r);
		if (!cells) continue;
		for (let c = footprint.start.col; c <= footprint.end.col; c++) {
			if (r === row && c === col) continue;
			const cell = cells.get(c);
			if (!cell) continue;
			if (cell.formula) return false;
			if (isSpilledCell(cell)) {
				const a = cell.spillAnchor;
				if (a.row === row && a.col === col) continue;
				if (claims(a, r, c)) return false;
				continue;
			}
			if (cell.value !== null) return false;
		}
	}
	return true;
}

/** Writes an array result into its footprint (the anchor cell gets the first element). */
export function writeFootprint(sheet: Worksheet, footprint: CellRange, m: Matrix): void {
	const { row, col } = footprint.start;
	for (let r = footprint.start.row; r <= footprint.end.row; r++) {
		for (let c = footprint.start.col; c <= footprint.end.col; c++) {
			const value = cellValueOf(m.get(r - row, c - col));
			const cell = getCell(sheet, r, c);
			if (r === row && c === col) {
				if (cell) cell.value = value;
				continue;
			}
			if (cell) {
				cell.value = value;
				(cell as SpilledCell).spillAnchor = { row, col };
			} else {
				const spilled: SpilledCell = { value, spillAnchor: { row, col } };
				putCell(sheet, r, c, spilled);
			}
		}
	}
}

/** Fills a CSE array-formula range (size-1 dimensions repeat, the rest is #N/A). */
export function fillArrayRange(sheet: Worksheet, range: CellRange, m: Matrix): void {
	const { row, col } = range.start;
	for (let r = range.start.row; r <= range.end.row; r++) {
		for (let c = range.start.col; c <= range.end.col; c++) {
			const value = cellValueOf(pick(m, r - row, c - col));
			const cell = getCell(sheet, r, c);
			if (cell) cell.value = value;
			else putCell(sheet, r, c, { value });
		}
	}
}

const cellValueOf = (value: CellValue): CellValue => (value === null ? 0 : value);
