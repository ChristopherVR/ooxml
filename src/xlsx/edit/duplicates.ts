import { type CellRange, normalizeRange, rangesIntersect } from '../address.js';
import { deleteCell, getCell, putCell } from '../cells.js';
import type { Cell } from '../model.js';
import { type EditContext, displayText, sheetAt } from './context.js';
import { translateFormula } from './deps.js';

export interface RemoveDuplicatesResult {
	/** Data rows removed. */
	removed: number;
	/** Unique data rows left (the header row is not counted). */
	remaining: number;
}

/**
 * Excel's Remove Duplicates: keeps the first data row of every distinct key and moves the
 * survivors up inside the range (cells outside it do not move). Keys compare the displayed text
 * of `columns` (absolute indices; empty means every column of the range) case-insensitively,
 * as Excel does (`3/8/2006` and `Mar 8, 2006` differ). Formulas of moved cells are adjusted
 * like a sort. One undo step.
 */
export function removeDuplicates(
	ctx: EditContext,
	s: number,
	range: CellRange,
	columns: number[],
	hasHeader: boolean,
): RemoveDuplicatesResult {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	const r = normalizeRange(range);
	const first = hasHeader ? r.start.row + 1 : r.start.row;
	if (first > r.end.row) return { removed: 0, remaining: 0 };
	const data: CellRange = { start: { row: first, col: r.start.col }, end: r.end };
	if (sheet.merges.some((m) => rangesIntersect(m, data)))
		throw new Error('Cannot remove duplicates from a range that contains merged cells.');
	const keyCols = columns.filter((c) => c >= r.start.col && c <= r.end.col);
	const cols = keyCols.length
		? keyCols
		: Array.from({ length: r.end.col - r.start.col + 1 }, (_v, i) => r.start.col + i);
	// Rows past the used area are all empty: stop at the last stored row.
	let lastStored = first;
	for (const row of sheet.rows.keys()) if (row > lastStored) lastStored = row;
	const lastRow = Math.min(r.end.row, lastStored);
	const keep: number[] = [];
	const seen = new Set<string>();
	for (let row = first; row <= lastRow; row++) {
		const key = cols
			.map((col) => displayText(workbook, getCell(sheet, row, col)).toLowerCase())
			.join('\u0000');
		if (seen.has(key)) continue;
		seen.add(key);
		keep.push(row);
	}
	const removed = lastRow - first + 1 - keep.length;
	if (!removed) return { removed: 0, remaining: keep.length };
	ctx.run(
		'Remove duplicates',
		'cells',
		[{ kind: 'cells', sheet: s, ranges: [data] }],
		() => {
			const moved = keep.map((row) => {
				const cells = new Map<number, Cell>();
				for (const [col, cell] of sheet.rows.get(row) ?? [])
					if (col >= r.start.col && col <= r.end.col) cells.set(col, cell);
				return { row, cells };
			});
			for (let row = first; row <= lastRow; row++)
				for (const col of [...(sheet.rows.get(row)?.keys() ?? [])])
					if (col >= r.start.col && col <= r.end.col) deleteCell(sheet, row, col);
			moved.forEach(({ row: from, cells }, i) => {
				const row = first + i;
				for (const [col, cell] of cells) {
					if (cell.formula !== undefined && row !== from)
						cell.formula = translateFormula(cell.formula, row - from, 0);
					putCell(sheet, row, col, cell);
				}
			});
		},
		{ sheet: s, ranges: [data] },
	);
	return { removed, remaining: keep.length };
}
