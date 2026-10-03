// Formatting whole rows and columns the way Excel 16 stores it. A cell with no record of its own
// takes its row's style, else its column's, so formatting a column over a row that has a style
// (or a row over styled columns) must create the cells where the two meet, or B3 would show only
// one of the two formats. Checked by saving from Excel: with row 3 filled yellow, bolding
// column B writes `<c r="B3" s=".."/>` with bold and yellow.
import { type CellRange, MAX_COL } from '../address.js';
import { getCell, putCell } from '../cells.js';
import type { Worksheet } from '../model.js';

type Patcher = (id: number | undefined) => number;

/** The sheet's columns as runs covering every column, with their style (0 when unstyled). */
function columnRuns(sheet: Worksheet): { min: number; max: number; style: number }[] {
	const runs: { min: number; max: number; style: number }[] = [];
	let next = 0;
	for (const c of [...sheet.columns].sort((a, b) => a.min - b.min)) {
		if (c.min > next) runs.push({ min: next, max: c.min - 1, style: 0 });
		runs.push({ min: c.min, max: c.max, style: c.styleId ?? 0 });
		next = c.max + 1;
	}
	if (next <= MAX_COL) runs.push({ min: next, max: MAX_COL, style: 0 });
	return runs;
}

/**
 * The style an unformatted row starts from. When one column entry covers most of the sheet
 * (as after formatting every column) Excel bases a new row style on it and only writes cells
 * where other columns differ; otherwise rows start from the default style.
 */
export function rowBaseStyle(sheet: Worksheet, row: number): number {
	const own = sheet.rowInfo.get(row)?.styleId;
	if (own !== undefined) return own;
	const wide = sheet.columns.find((c) => c.max - c.min + 1 > (MAX_COL + 1) / 2);
	return wide?.styleId ?? 0;
}

function addCell(sheet: Worksheet, row: number, col: number, styleId: number): void {
	putCell(sheet, row, col, styleId ? { value: null, styleId } : { value: null });
}

/**
 * Creates (with their current effective style, before patching) the cells where styled rows meet
 * the columns being formatted, wherever the patched result differs from what the cell would show
 * without a record of its own. Call before the columns are patched.
 */
export function addColumnIntersections(sheet: Worksheet, r: CellRange, patcher: Patcher): void {
	for (const [row, info] of sheet.rowInfo) {
		const rowStyle = info.styleId;
		if (rowStyle === undefined || row < r.start.row || row > r.end.row) continue;
		if (patcher(rowStyle) === rowStyle) continue;
		for (let col = r.start.col; col <= r.end.col; col++)
			if (!getCell(sheet, row, col)) addCell(sheet, row, col, rowStyle);
	}
}

/**
 * Creates the cells where unformatted rows being formatted meet styled columns whose patched
 * style differs from the new row style. Call before the rows are patched.
 */
export function addRowIntersections(sheet: Worksheet, r: CellRange, patcher: Patcher): void {
	const runs = columnRuns(sheet);
	for (let row = r.start.row; row <= r.end.row; row++) {
		if (sheet.rowInfo.get(row)?.styleId !== undefined) continue;
		const target = patcher(rowBaseStyle(sheet, row));
		for (const run of runs) {
			if (patcher(run.style) === target) continue;
			for (let col = run.min; col <= run.max; col++)
				if (!getCell(sheet, row, col)) addCell(sheet, row, col, run.style);
		}
	}
}
