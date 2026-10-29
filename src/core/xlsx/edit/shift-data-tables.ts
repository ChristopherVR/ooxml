// What-if data tables follow row and column inserts and deletes like formulas do.
//   - The result range `ref` is shifted like any range: an insert inside it widens it (the new
//     cells are recomputed from the new header cells), a delete trims it.
//   - The table is dropped (its results stay as plain values) when its whole input header row or
//     column is deleted, or when every result cell is.
//   - An input cell that is deleted keeps its text, gets `del1` / `del2` (Excel's marker) and the
//     table then shows `TABLE(#REF!,...)` and stops recalculating, as in Excel.
import { type CellAddress, type CellRange, formatAddress } from '../address';
import type { DataTable, Worksheet } from '../model';
import { parseInputCell } from './data-table';

type Move = (range: CellRange) => CellRange | undefined;

const point = (a: CellAddress): CellRange => ({ start: a, end: a });

function shiftInput(
	sheet: Worksheet,
	text: string,
	move: Move,
): { text: string; deleted: boolean } {
	const at = parseInputCell(sheet, text);
	if (!at) return { text, deleted: false };
	const moved = move(point(at));
	if (!moved) return { text, deleted: true };
	return { text: formatAddress(moved.start), deleted: false };
}

function shiftTable(sheet: Worksheet, table: DataTable, move: Move): DataTable | undefined {
	const ref = move(table.ref);
	if (!ref) return undefined;
	const { start, end } = table.ref;
	const header: CellRange[] = [
		{ start: { row: start.row - 1, col: start.col }, end: { row: start.row - 1, col: end.col } },
		{ start: { row: start.row, col: start.col - 1 }, end: { row: end.row, col: start.col - 1 } },
	];
	if (header.some((h) => h.start.row >= 0 && h.start.col >= 0 && !move(h))) return undefined;
	const next: DataTable = { ...table, ref };
	const first = shiftInput(sheet, table.r1, move);
	next.r1 = first.text;
	if (first.deleted || table.del1) next.del1 = true;
	if (table.r2 !== undefined) {
		const second = shiftInput(sheet, table.r2, move);
		next.r2 = second.text;
		if (second.deleted || table.del2) next.del2 = true;
	}
	return next;
}

/** Shifts the sheet's data tables for an axis edit (`move` shifts a range, undefined if deleted). */
export function shiftDataTables(sheet: Worksheet, move: Move): void {
	if (!sheet.dataTables?.length) return;
	const tables = sheet.dataTables.flatMap((t) => shiftTable(sheet, t, move) ?? []);
	if (tables.length) sheet.dataTables = tables;
	else delete sheet.dataTables;
}
