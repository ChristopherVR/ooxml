import { MAX_COL, type CellRange } from '../address.js';
import { forEachCellInRange, getCell, mergeAt } from '../cells.js';
import type { Workbook, Worksheet } from '../model.js';
import type { GridMetrics } from './metrics.js';
import type { HAlignView, MergeView } from './types.js';

/** How a cell takes part in a merge. */
export function mergeView(sheet: Worksheet, row: number, col: number): MergeView {
	const range = mergeAt(sheet, row, col);
	if (!range) return { anchor: false, hidden: false };
	const anchor = range.start.row === row && range.start.col === col;
	return { anchor, hidden: !anchor, range };
}

export interface SelectionStats {
	/** Non-empty cells (Excel's "Count"). */
	count: number;
	/** Numeric cells ("Numerical Count"). */
	numericCount: number;
	sum: number;
	average?: number;
	min?: number;
	max?: number;
}

/**
 * Status-bar statistics over selected ranges. Only stored cells are visited, so whole-column
 * selections stay cheap. Overlapping ranges count their shared cells once.
 */
export function selectionStats(
	workbook: Workbook,
	sheet: number | Worksheet,
	ranges: CellRange[],
): SelectionStats {
	const ws = typeof sheet === 'number' ? workbook.sheets[sheet] : sheet;
	const out: SelectionStats = { count: 0, numericCount: 0, sum: 0 };
	if (!ws) return out;
	const seen = new Set<string>();
	let min = Infinity;
	let max = -Infinity;
	for (const range of ranges)
		forEachCellInRange(ws, range, (cell, row, col) => {
			if (ranges.length > 1) {
				const key = `${row},${col}`;
				if (seen.has(key)) return;
				seen.add(key);
			}
			const value = cell.value;
			if (value === null || value === '') return;
			out.count++;
			if (typeof value === 'number' && Number.isFinite(value)) {
				out.numericCount++;
				out.sum += value;
				min = Math.min(min, value);
				max = Math.max(max, value);
			}
		});
	if (out.numericCount > 0) {
		out.average = out.sum / out.numericCount;
		out.min = min;
		out.max = max;
	}
	return out;
}

/** A cell that text may overflow into: nothing stored there, or an empty value without a formula. */
export function isOverflowTarget(sheet: Worksheet, row: number, col: number): boolean {
	const cell = getCell(sheet, row, col);
	if (cell && ((cell.value !== null && cell.value !== '') || cell.formula)) return false;
	return !mergeAt(sheet, row, col);
}

export interface OverflowExtent {
	startCol: number;
	endCol: number;
	/** Left and right edges of the painted text box in grid pixels. */
	left: number;
	right: number;
}

/**
 * How far a cell's text spills over empty neighbours. Left-aligned text extends right,
 * right-aligned text extends left and centred text both ways, each side stopping at the first
 * non-empty or merged cell. `textWidth` is the measured text width in pixels (at the metrics' zoom).
 */
export function overflowExtent(
	sheet: Worksheet,
	metrics: GridMetrics,
	row: number,
	col: number,
	textWidth: number,
	hAlign: HAlignView,
	maxSpan = 256,
): OverflowExtent {
	const width = metrics.colWidth(col);
	let startCol = col;
	let endCol = col;
	const result = (): OverflowExtent => ({
		startCol,
		endCol,
		left: metrics.colLeft(startCol),
		right: metrics.colLeft(endCol) + metrics.colWidth(endCol),
	});
	if (textWidth <= width || mergeAt(sheet, row, col)) return result();
	const extend = (need: number, step: 1 | -1): number => {
		let gained = 0;
		let c = col;
		for (let i = 0; i < maxSpan && gained < need; i++) {
			const next = c + step;
			if (next < 0 || next > MAX_COL || !isOverflowTarget(sheet, row, next)) break;
			c = next;
			gained += metrics.colWidth(c);
		}
		return c;
	};
	if (hAlign === 'left' || hAlign === 'justify' || hAlign === 'distributed')
		endCol = extend(textWidth - width, 1);
	else if (hAlign === 'right') startCol = extend(textWidth - width, -1);
	else if (hAlign === 'center' || hAlign === 'centerContinuous') {
		const side = (textWidth - width) / 2;
		startCol = extend(side, -1);
		endCol = extend(side, 1);
	}
	return result();
}
