// Sparse cell iteration helpers for the calc engine.
import type { CellRange } from '../address.js';
import type { Cell, Worksheet } from '../model.js';

/** Visits stored cells inside a range in row-major order, touching only cells that exist. */
export function scanRange(
	sheet: Worksheet,
	range: CellRange,
	visit: (cell: Cell, row: number, col: number) => void,
): void {
	const { start, end } = range;
	const rowCount = end.row - start.row + 1;
	const rows =
		rowCount <= sheet.rows.size * 2
			? null
			: [...sheet.rows.keys()].filter((r) => r >= start.row && r <= end.row).sort((a, b) => a - b);
	const visitRow = (r: number): void => {
		const cells = sheet.rows.get(r);
		if (!cells || cells.size === 0) return;
		const colCount = end.col - start.col + 1;
		if (colCount <= cells.size * 2) {
			for (let c = start.col; c <= end.col; c++) {
				const cell = cells.get(c);
				if (cell) visit(cell, r, c);
			}
			return;
		}
		const cols = [...cells.keys()]
			.filter((c) => c >= start.col && c <= end.col)
			.sort((a, b) => a - b);
		for (const c of cols) {
			const cell = cells.get(c);
			if (cell) visit(cell, r, c);
		}
	};
	if (rows) for (const r of rows) visitRow(r);
	else for (let r = start.row; r <= end.row; r++) visitRow(r);
}

/** One past the last stored row and column. */
export function sheetBounds(sheet: Worksheet): { rows: number; cols: number } {
	let rows = 0;
	let cols = 0;
	for (const [r, cells] of sheet.rows) {
		if (cells.size === 0) continue;
		rows = Math.max(rows, r + 1);
		for (const c of cells.keys()) cols = Math.max(cols, c + 1);
	}
	return { rows, cols };
}
