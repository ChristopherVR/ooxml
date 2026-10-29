// Scenario Manager scenarios follow row and column inserts and deletes like formulas do.
//   - A changing cell moves with its cell. When it is deleted the scenario keeps the entry with
//     `deleted` set (Excel's marker), so showing the scenario skips it.
//   - The result cells (`sqref`) are shifted like any range list; deleted ranges drop out.
import { type CellRange, formatAddress, formatRange, parseAddress, parseRange } from '../address';
import type { Worksheet } from '../model';

type Move = (range: CellRange) => CellRange | undefined;

const isSingle = (r: CellRange): boolean => r.start.row === r.end.row && r.start.col === r.end.col;

function shiftSqref(sqref: string, move: Move): string | undefined {
	const parts = sqref.split(/\s+/).flatMap((part) => {
		const range = parseRange(part.replace(/\$/g, ''));
		if (!range) return [part];
		const moved = move(range);
		return moved ? [isSingle(moved) ? formatAddress(moved.start) : formatRange(moved)] : [];
	});
	return parts.length ? parts.join(' ') : undefined;
}

/** Shifts the sheet's scenarios for an axis edit (`move` shifts a range, undefined if deleted). */
export function shiftScenarios(sheet: Worksheet, move: Move): void {
	const set = sheet.scenarios;
	if (!set) return;
	for (const scenario of set.list)
		for (const cell of scenario.cells) {
			if (cell.deleted) continue;
			const at = parseAddress(cell.ref.replace(/\$/g, ''));
			if (!at) continue;
			const moved = move({ start: at, end: at });
			if (moved) cell.ref = formatAddress(moved.start);
			else cell.deleted = true;
		}
	if (set.sqref) {
		const next = shiftSqref(set.sqref, move);
		if (next) set.sqref = next;
		else delete set.sqref;
	}
}
