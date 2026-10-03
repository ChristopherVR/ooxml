// Pure selection arithmetic for the grid (no DOM): Excel-style extension, header selections,
// moving the active cell inside a selection (Enter/Tab) and merge expansion.
import {
	MAX_COL,
	MAX_ROW,
	normalizeRange,
	rangeContains,
	type CellAddress,
	type CellRange,
	type Worksheet,
} from '@christophervr/xlsx-core';
import type { Selection } from '../context.js';
import { isWholeCols, isWholeRows } from '../selection.js';

export const cellRange = (at: CellAddress): CellRange => ({ start: { ...at }, end: { ...at } });

export const spanRange = (a: CellAddress, b: CellAddress): CellRange =>
	normalizeRange({ start: { ...a }, end: { ...b } });

export const rowsRange = (r1: number, r2: number): CellRange =>
	spanRange({ row: r1, col: 0 }, { row: r2, col: MAX_COL });

export const colsRange = (c1: number, c2: number): CellRange =>
	spanRange({ row: 0, col: c1 }, { row: MAX_ROW, col: c2 });

export const ALL_RANGE: CellRange = {
	start: { row: 0, col: 0 },
	end: { row: MAX_ROW, col: MAX_COL },
};

// One A1 formatter for the whole editor (events, getSelection, announcements): see selection.ts.
export { isWholeCols, isWholeRows, rangeRef, selectionRef } from '../selection.js';

export const isSingleCell = (range: CellRange): boolean =>
	range.start.row === range.end.row && range.start.col === range.end.col;

/** Grows a range until no merge sticks out of it (Excel never selects part of a merge). */
export function expandForMerges(sheet: Worksheet, range: CellRange): CellRange {
	let out = normalizeRange(range);
	if (!sheet.merges.length || isWholeRows(out) || isWholeCols(out)) return out;
	for (let pass = 0; pass < 8; pass++) {
		let grown = false;
		for (const merge of sheet.merges) {
			const hit =
				merge.start.row <= out.end.row &&
				merge.end.row >= out.start.row &&
				merge.start.col <= out.end.col &&
				merge.end.col >= out.start.col;
			if (!hit) continue;
			const next = {
				start: {
					row: Math.min(out.start.row, merge.start.row),
					col: Math.min(out.start.col, merge.start.col),
				},
				end: {
					row: Math.max(out.end.row, merge.end.row),
					col: Math.max(out.end.col, merge.end.col),
				},
			};
			if (
				next.start.row !== out.start.row ||
				next.start.col !== out.start.col ||
				next.end.row !== out.end.row ||
				next.end.col !== out.end.col
			) {
				out = next;
				grown = true;
			}
		}
		if (!grown) break;
	}
	return out;
}

/** The corner of the last range that Shift+arrows move (opposite the anchor). */
export function movingEnd(selection: Selection): CellAddress {
	const range = selection.ranges[selection.ranges.length - 1] ?? cellRange(selection.active);
	const { anchor } = selection;
	return {
		row: range.start.row === anchor.row ? range.end.row : range.start.row,
		col: range.start.col === anchor.col ? range.end.col : range.start.col,
	};
}

/** A fresh one-cell selection at `at` (merge-expanded). */
export function selectCell(
	sheetIndex: number,
	sheet: Worksheet | undefined,
	at: CellAddress,
): Selection {
	const range = sheet ? expandForMerges(sheet, cellRange(at)) : cellRange(at);
	return { sheet: sheetIndex, active: { ...at }, anchor: { ...at }, ranges: [range] };
}

/** Extends the last range from the anchor to `to` (Shift+click, Shift+arrows, drag). */
export function extendSelection(
	selection: Selection,
	sheet: Worksheet | undefined,
	to: CellAddress,
): Selection {
	const range = spanRange(selection.anchor, to);
	const last = sheet ? expandForMerges(sheet, range) : range;
	return { ...selection, ranges: [...selection.ranges.slice(0, -1), last] };
}

/** Ctrl+click: adds a new range and makes its cell active. */
export function addRange(
	selection: Selection,
	sheet: Worksheet | undefined,
	at: CellAddress,
): Selection {
	const range = sheet ? expandForMerges(sheet, cellRange(at)) : cellRange(at);
	return {
		...selection,
		active: { ...at },
		anchor: { ...at },
		ranges: [...selection.ranges, range],
	};
}

/** Every cell of the selection counted once is too much for whole columns; this is the cell count. */
export function selectionCellCount(selection: Selection): number {
	return selection.ranges.reduce(
		(sum, r) => sum + (r.end.row - r.start.row + 1) * (r.end.col - r.start.col + 1),
		0,
	);
}

/**
 * Enter / Tab inside a multi-cell selection: the active cell walks the selection (rows first for
 * Tab, columns first for Enter), wrapping to the next range, and the ranges stay.
 */
export function moveWithinSelection(
	selection: Selection,
	direction: 'down' | 'up' | 'right' | 'left',
): Selection | undefined {
	if (
		selection.ranges.length === 1 &&
		isSingleCell(selection.ranges[0] ?? cellRange(selection.active))
	)
		return undefined;
	if (selectionCellCount(selection) <= 1) return undefined;
	const forward = direction === 'down' || direction === 'right';
	const byRows = direction === 'right' || direction === 'left';
	let index = selection.ranges.findIndex((r) => rangeContains(r, selection.active));
	if (index < 0) index = selection.ranges.length - 1;
	let range = selection.ranges[index] ?? cellRange(selection.active);
	let { row, col } = selection.active;
	const step = (): boolean => {
		if (byRows) {
			col += forward ? 1 : -1;
			if (col > range.end.col || col < range.start.col) {
				col = forward ? range.start.col : range.end.col;
				row += forward ? 1 : -1;
				if (row > range.end.row || row < range.start.row) return false;
			}
		} else {
			row += forward ? 1 : -1;
			if (row > range.end.row || row < range.start.row) {
				row = forward ? range.start.row : range.end.row;
				col += forward ? 1 : -1;
				if (col > range.end.col || col < range.start.col) return false;
			}
		}
		return true;
	};
	if (!step()) {
		index = (index + (forward ? 1 : -1) + selection.ranges.length) % selection.ranges.length;
		range = selection.ranges[index] ?? range;
		row = forward ? range.start.row : range.end.row;
		col = forward ? range.start.col : range.end.col;
	}
	return { ...selection, active: { row, col } };
}

/** Clamps an address to the sheet. */
export const clampAddress = (at: CellAddress): CellAddress => ({
	row: Math.max(0, Math.min(MAX_ROW, at.row)),
	col: Math.max(0, Math.min(MAX_COL, at.col)),
});

/** Whether a row / column is part of the selection (for header highlighting). */
export function rowSelected(selection: Selection, row: number): 'none' | 'part' | 'full' {
	let state: 'none' | 'part' | 'full' = 'none';
	for (const r of selection.ranges)
		if (row >= r.start.row && row <= r.end.row) {
			if (isWholeRows(r)) return 'full';
			state = 'part';
		}
	return state;
}

export function colSelected(selection: Selection, col: number): 'none' | 'part' | 'full' {
	let state: 'none' | 'part' | 'full' = 'none';
	for (const r of selection.ranges)
		if (col >= r.start.col && col <= r.end.col) {
			if (isWholeCols(r)) return 'full';
			state = 'part';
		}
	return state;
}
