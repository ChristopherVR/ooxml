import type { CellAddress, CellRange } from './address.js';
import type { Cell, CellValue, Worksheet } from './model.js';

/** The cell at a position, or `undefined` for an empty one. */
export const getCell = (sheet: Worksheet, row: number, col: number): Cell | undefined =>
	sheet.rows.get(row)?.get(col);

export const getCellAt = (sheet: Worksheet, address: CellAddress): Cell | undefined =>
	getCell(sheet, address.row, address.col);

/** A cell's value; empty cells are `null`. */
export const getValue = (sheet: Worksheet, row: number, col: number): CellValue =>
	getCell(sheet, row, col)?.value ?? null;

/** Stores `cell` at a position, replacing what was there. */
export function putCell(sheet: Worksheet, row: number, col: number, cell: Cell): void {
	let cells = sheet.rows.get(row);
	if (!cells) {
		cells = new Map();
		sheet.rows.set(row, cells);
	}
	cells.set(col, cell);
}

/** Removes the cell at a position; empty rows are dropped from the map. */
export function deleteCell(sheet: Worksheet, row: number, col: number): void {
	const cells = sheet.rows.get(row);
	if (!cells) return;
	cells.delete(col);
	if (cells.size === 0) sheet.rows.delete(row);
}

/** Whether a cell holds nothing worth saving (no value, no formula, default style). */
export const isEmptyCell = (cell: Cell | undefined): boolean =>
	!cell ||
	((cell.value === null || cell.value === '') && !cell.formula && !cell.styleId && !cell.richText);

/** Visits every stored cell in row-major order. */
export function forEachCell(
	sheet: Worksheet,
	visit: (cell: Cell, row: number, col: number) => void,
): void {
	for (const row of [...sheet.rows.keys()].sort((a, b) => a - b)) {
		const cells = sheet.rows.get(row);
		if (!cells) continue;
		for (const col of [...cells.keys()].sort((a, b) => a - b)) {
			const cell = cells.get(col);
			if (cell) visit(cell, row, col);
		}
	}
}

/** Visits stored cells inside a range, in row-major order. */
export function forEachCellInRange(
	sheet: Worksheet,
	range: CellRange,
	visit: (cell: Cell, row: number, col: number) => void,
): void {
	const rows = [...sheet.rows.keys()]
		.filter((row) => row >= range.start.row && row <= range.end.row)
		.sort((a, b) => a - b);
	for (const row of rows) {
		const cells = sheet.rows.get(row);
		if (!cells) continue;
		const cols = [...cells.keys()]
			.filter((col) => col >= range.start.col && col <= range.end.col)
			.sort((a, b) => a - b);
		for (const col of cols) {
			const cell = cells.get(col);
			if (cell) visit(cell, row, col);
		}
	}
}

/**
 * The used range: the smallest range holding every stored cell (`undefined` for an empty sheet).
 * Styled-but-empty cells count, matching Excel's `dimension`.
 */
export function usedRange(sheet: Worksheet): CellRange | undefined {
	let minRow = Infinity;
	let maxRow = -1;
	let minCol = Infinity;
	let maxCol = -1;
	for (const [row, cells] of sheet.rows) {
		if (cells.size === 0) continue;
		minRow = Math.min(minRow, row);
		maxRow = Math.max(maxRow, row);
		for (const col of cells.keys()) {
			minCol = Math.min(minCol, col);
			maxCol = Math.max(maxCol, col);
		}
	}
	if (maxRow < 0) return undefined;
	return { start: { row: minRow, col: minCol }, end: { row: maxRow, col: maxCol } };
}

/** The merged range a cell belongs to, if any. */
export function mergeAt(sheet: Worksheet, row: number, col: number): CellRange | undefined {
	return sheet.merges.find(
		(m) => row >= m.start.row && row <= m.end.row && col >= m.start.col && col <= m.end.col,
	);
}
