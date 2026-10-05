import { MAX_COL, MAX_ROW, type CellAddress } from '../address.js';
import { usedRange } from '../cells.js';
import type { Worksheet } from '../model.js';

export type NavigationKey =
	| 'up'
	| 'down'
	| 'left'
	| 'right'
	| 'home'
	| 'end'
	| 'pageUp'
	| 'pageDown'
	| 'ctrlUp'
	| 'ctrlDown'
	| 'ctrlLeft'
	| 'ctrlRight'
	| 'ctrlHome'
	| 'ctrlEnd';

const rowHidden = (sheet: Worksheet, row: number): boolean => {
	const info = sheet.rowInfo.get(row);
	return info?.hidden === true || (info?.height !== undefined && info.height <= 0);
};

const colHidden = (sheet: Worksheet, col: number): boolean =>
	sheet.columns.some(
		(c) =>
			col >= c.min &&
			col <= c.max &&
			(c.hidden === true || (c.width !== undefined && c.width <= 0)),
	);

const filled = (sheet: Worksheet, row: number, col: number): boolean => {
	const cell = sheet.rows.get(row)?.get(col);
	return (
		cell !== undefined && ((cell.value !== null && cell.value !== '') || cell.formula !== undefined)
	);
};

/** One visible step from `index` along an axis, or `index` when there is none. */
function step(index: number, dir: 1 | -1, max: number, hidden: (i: number) => boolean): number {
	for (let i = index + dir; i >= 0 && i <= max; i += dir) if (!hidden(i)) return i;
	return index;
}

/** The first / last visible index from an end of the axis. */
function edge(dir: 1 | -1, max: number, hidden: (i: number) => boolean): number {
	const start = dir > 0 ? max : 0;
	if (!hidden(start)) return start;
	return step(start, (dir > 0 ? -1 : 1) as 1 | -1, max, hidden);
}

/**
 * Excel's Ctrl+arrow along one axis: inside a block of filled cells, jump to the block's last
 * cell; otherwise jump to the next filled cell, or to the sheet edge when there is none. Hidden
 * cells are skipped as if absent. `filledAt` lists the visible filled indices on the line.
 */
function jump(
	index: number,
	dir: 1 | -1,
	max: number,
	hidden: (i: number) => boolean,
	filledSet: Set<number>,
): number {
	const next = step(index, dir, max, hidden);
	if (next === index) return index;
	if (filledSet.has(index) && filledSet.has(next)) {
		let at = next;
		for (;;) {
			const after = step(at, dir, max, hidden);
			if (after === at || !filledSet.has(after)) return at;
			at = after;
		}
	}
	const candidates = [...filledSet].filter((i) => (dir > 0 ? i > index : i < index) && !hidden(i));
	if (!candidates.length) return edge(dir, max, hidden);
	return dir > 0 ? Math.min(...candidates) : Math.max(...candidates);
}

function filledInColumn(sheet: Worksheet, col: number): Set<number> {
	const out = new Set<number>();
	for (const row of sheet.rows.keys()) if (filled(sheet, row, col)) out.add(row);
	return out;
}

function filledInRow(sheet: Worksheet, row: number): Set<number> {
	const out = new Set<number>();
	const cells = sheet.rows.get(row);
	if (cells) for (const col of cells.keys()) if (filled(sheet, row, col)) out.add(col);
	return out;
}

/** The cell a navigation key moves the active cell to (Excel semantics, hidden rows/cols skipped). */
export function navigate(
	sheet: Worksheet,
	from: CellAddress,
	key: NavigationKey,
	pageRows = 20,
): CellAddress {
	const isRowHidden = (r: number): boolean => rowHidden(sheet, r);
	const isColHidden = (c: number): boolean => colHidden(sheet, c);
	const { row, col } = from;
	switch (key) {
		case 'up':
			return { row: step(row, -1, MAX_ROW, isRowHidden), col };
		case 'down':
			return { row: step(row, 1, MAX_ROW, isRowHidden), col };
		case 'left':
			return { row, col: step(col, -1, MAX_COL, isColHidden) };
		case 'right':
			return { row, col: step(col, 1, MAX_COL, isColHidden) };
		case 'pageUp':
		case 'pageDown': {
			const dir = key === 'pageUp' ? -1 : 1;
			let r = row;
			for (let i = 0; i < Math.max(1, pageRows); i++) {
				const next = step(r, dir, MAX_ROW, isRowHidden);
				if (next === r) break;
				r = next;
			}
			return { row: r, col };
		}
		case 'home': {
			const first = isColHidden(0) ? step(0, 1, MAX_COL, isColHidden) : 0;
			return { row, col: first };
		}
		case 'end': {
			// Last filled cell in the row (Excel's End, Right from the row's start).
			const cols = [...filledInRow(sheet, row)].filter((c) => !isColHidden(c));
			return { row, col: cols.length ? Math.max(...cols) : col };
		}
		case 'ctrlUp':
			return { row: jump(row, -1, MAX_ROW, isRowHidden, filledInColumn(sheet, col)), col };
		case 'ctrlDown':
			return { row: jump(row, 1, MAX_ROW, isRowHidden, filledInColumn(sheet, col)), col };
		case 'ctrlLeft':
			return { row, col: jump(col, -1, MAX_COL, isColHidden, filledInRow(sheet, row)) };
		case 'ctrlRight':
			return { row, col: jump(col, 1, MAX_COL, isColHidden, filledInRow(sheet, row)) };
		case 'ctrlHome': {
			// The first cell of the scrolling pane (below and right of frozen panes).
			const r0 = sheet.view.freeze?.rows ?? 0;
			const c0 = sheet.view.freeze?.cols ?? 0;
			return {
				row: isRowHidden(r0) ? step(r0, 1, MAX_ROW, isRowHidden) : r0,
				col: isColHidden(c0) ? step(c0, 1, MAX_COL, isColHidden) : c0,
			};
		}
		case 'ctrlEnd': {
			const used = usedRange(sheet);
			if (!used) return { row: 0, col: 0 };
			const r = isRowHidden(used.end.row)
				? step(used.end.row, -1, MAX_ROW, isRowHidden)
				: used.end.row;
			const c = isColHidden(used.end.col)
				? step(used.end.col, -1, MAX_COL, isColHidden)
				: used.end.col;
			return { row: r, col: c };
		}
	}
}
