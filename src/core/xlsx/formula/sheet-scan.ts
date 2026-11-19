import type { Worksheet } from '../model';

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
