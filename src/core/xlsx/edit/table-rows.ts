// Row helpers for table edits: empty checks, clearing, and the totals row cells.
import { getCell } from '../cells.js';
import type { Table, Worksheet } from '../model.js';
import { clearContents, writeValue } from './cell-values.js';
import { ensureCell, pruneCell } from './context.js';

/** Whether a row holds no values or formulas between two columns. */
export function rowIsEmpty(sheet: Worksheet, row: number, lo: number, hi: number): boolean {
	const cells = sheet.rows.get(row);
	if (!cells) return true;
	for (const [col, cell] of cells)
		if (
			col >= lo &&
			col <= hi &&
			(cell.formula !== undefined || (cell.value !== null && cell.value !== ''))
		)
			return false;
	return true;
}

/** Clears the contents of a row between two columns (formats stay). */
export function clearRow(sheet: Worksheet, row: number, lo: number, hi: number): void {
	for (let col = lo; col <= hi; col++) {
		const cell = getCell(sheet, row, col);
		if (!cell) continue;
		clearContents(cell);
		pruneCell(sheet, row, col);
	}
}

const SUBTOTAL: Record<string, number> = {
	average: 101,
	countNums: 102,
	count: 103,
	max: 104,
	min: 105,
	stdDev: 107,
	sum: 109,
	var: 110,
};

const escapeColumn = (name: string): string => name.replace(/['#[\]]/g, "'$&");

/**
 * Fills the totals row from the columns' `totalsRowLabel` / `totalsRowFunction`. When no column
 * has either yet, Excel's default applies: `Total` in the first column and a sum in the last.
 */
export function applyTotalsRow(sheet: Worksheet, table: Table): void {
	const row = table.range.end.row;
	const { columns } = table;
	if (!columns.some((c) => c.totalsRowFunction || c.totalsRowLabel)) {
		const first = columns[0];
		const last = columns[columns.length - 1];
		if (first && columns.length > 1) first.totalsRowLabel = 'Total';
		if (last) last.totalsRowFunction = 'sum';
	}
	columns.forEach((column, i) => {
		const col = table.range.start.col + i;
		const code = SUBTOTAL[column.totalsRowFunction ?? ''];
		if (code !== undefined) {
			const cell = ensureCell(sheet, row, col);
			clearContents(cell);
			cell.formula = `SUBTOTAL(${code},${table.name}[${escapeColumn(column.name)}])`;
		} else if (column.totalsRowLabel && column.totalsRowFunction !== 'custom')
			writeValue(sheet, row, col, column.totalsRowLabel);
	});
}
