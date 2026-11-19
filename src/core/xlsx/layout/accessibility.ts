// What the grid exposes to assistive technology, DOM-free: which cells the accessible grid holds
// (the painted window plus the active cell, merges as spans) and what one cell announces besides
// its shown text (its table column, formula and comment). The editor words and renders it.
import { cellKey, formatAddress, type CellAddress } from '../address';
import { getCell } from '../cells';
import type { Worksheet } from '../model';
import type { VisibleCells } from './viewport';

export interface CellDescription {
	/** The cell's A1 reference. */
	ref: string;
	/** The name of the table column the cell sits in, for cells below a table's header row. */
	header?: string;
	/** The formula as Excel shows it in the formula bar (`=A1*2`, `{=A1:A3*2}` for array formulas). */
	formula?: string;
	/** The text of the cell's comment or note. */
	comment?: string;
}

export function describeCell(sheet: Worksheet, at: CellAddress): CellDescription {
	const out: CellDescription = { ref: formatAddress(at) };
	const table = sheet.tables.find(
		(t) =>
			t.headerRow &&
			at.row > t.range.start.row &&
			at.row <= t.range.end.row &&
			at.col >= t.range.start.col &&
			at.col <= t.range.end.col,
	);
	const header = table?.columns[at.col - (table?.range.start.col ?? 0)]?.name;
	if (header) out.header = header;
	const cell = getCell(sheet, at.row, at.col);
	if (cell?.formula !== undefined)
		out.formula = cell.arrayRange ? `{=${cell.formula}}` : `=${cell.formula}`;
	const comment = sheet.comments.find((c) => c.address.row === at.row && c.address.col === at.col);
	if (comment?.text) out.comment = comment.text;
	return out;
}

export interface AccessibleWindow {
	/** Rows of the accessible grid, ascending and unique. */
	rows: number[];
	/** Columns of the accessible grid, ascending and unique. */
	cols: number[];
	/** Merge anchors in the window (by {@link cellKey}) with the rows and columns they span. */
	spans: Map<number, { rows: number; cols: number }>;
	/** Cells a merge covers without anchoring it: they have no grid cell of their own. */
	covered: Set<number>;
}

const union = (a: readonly number[], b: readonly number[], extra?: number): number[] => {
	const all = new Set([...a, ...b]);
	if (extra !== undefined) all.add(extra);
	return [...all].sort((x, y) => x - y);
};

/**
 * The cells the accessible grid holds: every painted row and column (frozen and scrolling panes)
 * and the active cell's row and column, so the active cell always has a grid cell to point at.
 */
export function accessibleWindow(
	sheet: Worksheet | undefined,
	visible: VisibleCells,
	active?: CellAddress,
): AccessibleWindow {
	const rows = union(visible.frozenRows, visible.rows, active?.row);
	const cols = union(visible.frozenCols, visible.cols, active?.col);
	const spans = new Map<number, { rows: number; cols: number }>();
	const covered = new Set<number>();
	const first = { row: rows[0] ?? 0, col: cols[0] ?? 0 };
	const last = { row: rows.at(-1) ?? -1, col: cols.at(-1) ?? -1 };
	for (const m of sheet?.merges ?? []) {
		if (m.end.row < first.row || m.start.row > last.row) continue;
		if (m.end.col < first.col || m.start.col > last.col) continue;
		spans.set(cellKey(m.start.row, m.start.col), {
			rows: m.end.row - m.start.row + 1,
			cols: m.end.col - m.start.col + 1,
		});
		for (const r of rows) {
			if (r < m.start.row || r > m.end.row) continue;
			for (const c of cols)
				if (c >= m.start.col && c <= m.end.col && (r !== m.start.row || c !== m.start.col))
					covered.add(cellKey(r, c));
		}
	}
	return { rows, cols, spans, covered };
}
