// The order Find Next and Find Previous walk matches in.
import type { FindMatch } from '../index';

/** Matches after the active cell first (wrapping), as Find Next walks them. */
export function orderFrom(
	matches: FindMatch[],
	sheet: number,
	row: number,
	col: number,
	byColumns: boolean,
): FindMatch[] {
	const after = (m: FindMatch): boolean =>
		m.sheet > sheet ||
		(m.sheet === sheet &&
			(byColumns
				? m.col > col || (m.col === col && m.row > row)
				: m.row > row || (m.row === row && m.col > col)));
	const index = matches.findIndex(after);
	return index < 0 ? matches : [...matches.slice(index), ...matches.slice(0, index)];
}

/** The match Find Next / Find Previous lands on from the active cell, wrapping around. */
export function stepMatch(
	matches: FindMatch[],
	sheet: number,
	row: number,
	col: number,
	byColumns: boolean,
	backwards = false,
): FindMatch | undefined {
	const list = orderFrom(matches, sheet, row, col, byColumns);
	if (!backwards) return list[0];
	// `list` starts after the active cell and ends at or before it: the last entry is the previous.
	const last = list[list.length - 1];
	const atActive = last && last.sheet === sheet && last.row === row && last.col === col;
	return atActive && list.length > 1 ? list[list.length - 2] : last;
}
