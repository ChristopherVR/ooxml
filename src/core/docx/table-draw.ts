// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Word's freeform Draw Table and its structural Eraser, on the table model. A line drawn inside a
// cell splits it (a vertical line at the drawn position, adding a grid column when no column
// boundary is near; a horizontal line into two rows, or a vertical merge at the nearest row
// boundary), and erasing a line between two cells merges them when together they form a
// rectangle. The editor measures the pointer against the cell it hit and passes the position as a
// fraction of that cell; everything here works on the grid (`w:tblGrid`, `gridSpan`, `vMerge`).
import type { Table } from './model';
import { continuationsOf, mergeOrigin, rowSpanOf, tableGrid, type GridCell } from './table-grid';
import { copyTable, emptyCell, type IdFactory } from './table-structure-cells';
import { mergeCells, splitCell, type GridRect } from './table-structure-merge';
import { normalizeMerges } from './table-structure';
import type { Twips } from './units';

/** A cell of `Table.rows`: its row and its position in that row. */
export interface DrawCellRef {
	row: number;
	index: number;
}

/** A line drawn across one cell, at `fraction` (0 to 1) of its width or height. */
export interface DrawnLine extends DrawCellRef {
	orientation: 'vertical' | 'horizontal';
	fraction: number;
}

/** A cell side the Eraser touched; `fraction` is where along that side (0 to 1). */
export interface ErasedEdge extends DrawCellRef {
	side: 'top' | 'bottom' | 'left' | 'right';
	fraction?: number;
}

export type EraseOutcome =
	| { kind: 'merged'; table: Table }
	/** The side is on the table's outside: the editor removes that border line instead. */
	| { kind: 'outer' }
	/** The two cells beside the line do not make a rectangle (or cannot be merged). */
	| { kind: 'refused' };

export interface DrawOptions {
	/** Total width the grid takes when the table has no `w:tblGrid` (columns are then even). */
	defaultWidthTwips?: number;
}

/** Closer than this to an existing column boundary, a vertical line snaps to it (0.05 inch). */
export const DRAW_SNAP_TWIPS = 72;
/** The narrowest cell a drawn line may leave on either side (0.08 inch). */
export const DRAW_MIN_CELL_TWIPS = 115;
/** The smallest table a drawn rectangle makes (0.1 inch square). */
export const DRAW_MIN_TABLE_TWIPS = 144;

const clampFraction = (value: number): number =>
	Number.isFinite(value) ? Math.min(1, Math.max(0, value)) : 0.5;

function widthsOf(table: Table, columns: number, options: DrawOptions): number[] | undefined {
	if (table.grid?.length === columns) return table.grid.map(Number);
	const total = options.defaultWidthTwips;
	if (!total || !Number.isFinite(total) || total <= 0) return undefined;
	const each = Math.floor(total / columns);
	return Array.from({ length: columns }, (_, i) =>
		i === columns - 1 ? total - each * (columns - 1) : each,
	);
}

const sum = (values: number[], from: number, to: number): number =>
	values.slice(from, to).reduce((total, value) => total + value, 0);

/** The size, in twips, of the one-cell table a rectangle drawn in the body makes. */
export function drawnTableBox(
	widthTwips: number,
	heightTwips: number,
	maxWidthTwips?: number,
): { widthTwips: Twips; heightTwips: Twips } | undefined {
	if (!Number.isFinite(widthTwips) || !Number.isFinite(heightTwips)) return undefined;
	if (widthTwips < DRAW_MIN_TABLE_TWIPS || heightTwips < DRAW_MIN_TABLE_TWIPS) return undefined;
	const width =
		maxWidthTwips && maxWidthTwips > 0 ? Math.min(widthTwips, maxWidthTwips) : widthTwips;
	return { widthTwips: Math.round(width) as Twips, heightTwips: Math.round(heightTwips) as Twips };
}

/**
 * Splits the cell a line was drawn across. Returns the edited table, or undefined when the line
 * lands too near the cell's edge or the table is not a consistent grid.
 */
export function drawTableLine(
	table: Table,
	line: DrawnLine,
	newId: IdFactory,
	options: DrawOptions = {},
): Table | undefined {
	const grid = tableGrid(table);
	const entry = grid?.rows[line.row]?.[line.index];
	if (!grid || !entry) return undefined;
	const origin = mergeOrigin(grid, entry);
	return line.orientation === 'horizontal'
		? splitAcross(table, grid, origin, clampFraction(line.fraction), newId)
		: splitDown(table, grid, origin, clampFraction(line.fraction), newId, options);
}

function splitAcross(
	table: Table,
	grid: NonNullable<ReturnType<typeof tableGrid>>,
	origin: GridCell,
	fraction: number,
	newId: IdFactory,
): Table | undefined {
	const height = rowSpanOf(grid, origin);
	if (height === 1) return splitCell(table, origin.row, origin.index, 2, 1, newId);
	// A vertical merge comes apart at the row boundary nearest the line.
	const boundary = Math.min(height - 1, Math.max(1, Math.round(fraction * height)));
	const below = continuationsOf(grid, origin)[boundary - 1];
	if (!below) return undefined;
	const next = copyTable(table);
	const cell = next.rows[below.row]?.[below.index];
	if (!cell) return undefined;
	cell.verticalMerge = 'restart';
	normalizeMerges(next);
	return next;
}

function splitDown(
	table: Table,
	grid: NonNullable<ReturnType<typeof tableGrid>>,
	origin: GridCell,
	fraction: number,
	newId: IdFactory,
	options: DrawOptions,
): Table | undefined {
	const widths = widthsOf(table, grid.columns, options);
	if (!widths) return undefined;
	const left = sum(widths, 0, origin.col);
	const width = sum(widths, origin.col, origin.col + origin.span);
	const x = left + fraction * width;
	if (x - left < DRAW_MIN_CELL_TWIPS || left + width - x < DRAW_MIN_CELL_TWIPS) return undefined;
	const next = copyTable(table);
	next.grid = widths.map((w) => w as Twips);
	// An existing boundary inside the cell near the line is used as it is.
	let boundary: number | undefined;
	let best = DRAW_SNAP_TWIPS + 1;
	for (let k = origin.col + 1; k < origin.col + origin.span; k++) {
		const distance = Math.abs(sum(widths, 0, k) - x);
		if (distance < best) [boundary, best] = [k, distance];
	}
	let span = origin.span;
	if (boundary === undefined) {
		// Cut the grid column under the line in two; every cell over it widens to keep covering it.
		let column = origin.col;
		while (column < origin.col + origin.span - 1 && sum(widths, 0, column + 1) <= x) column++;
		const start = sum(widths, 0, column);
		const cut = Math.round(x - start);
		const rest = (widths[column] ?? 0) - cut;
		if (cut <= 0 || rest <= 0) return undefined;
		next.grid.splice(column, 1, cut as Twips, rest as Twips);
		grid.rows.forEach((cells, r) => {
			const over = cells.find((e) => e.col <= column && column < e.col + e.span);
			const cell = over && next.rows[r]?.[over.index];
			if (cell) cell.gridSpan = (cell.gridSpan ?? 1) + 1;
		});
		boundary = column + 1;
		span += 1;
	}
	const leftSpan = boundary - origin.col;
	const rightSpan = span - leftSpan;
	const gridWidths = next.grid.map(Number);
	const leftWidth = sum(gridWidths, origin.col, boundary) as Twips;
	const rightWidth = sum(gridWidths, boundary, origin.col + span) as Twips;
	for (const member of [origin, ...continuationsOf(grid, origin)]) {
		const row = next.rows[member.row];
		const cell = row?.[member.index];
		if (!row || !cell) return undefined;
		const right = emptyCell(
			newId,
			cell,
			rightSpan,
			cell.widthTwips === undefined ? undefined : rightWidth,
		);
		if (cell.verticalMerge) right.verticalMerge = cell.verticalMerge;
		if (leftSpan > 1) cell.gridSpan = leftSpan;
		else delete cell.gridSpan;
		if (cell.widthTwips !== undefined) cell.widthTwips = leftWidth;
		row.splice(member.index + 1, 0, right);
	}
	return tableGrid(next) ? next : undefined;
}

const rectOf = (grid: NonNullable<ReturnType<typeof tableGrid>>, cell: GridCell): GridRect => ({
	firstRow: cell.row,
	lastRow: cell.row + rowSpanOf(grid, cell) - 1,
	firstCol: cell.col,
	lastCol: cell.col + cell.span - 1,
});

/**
 * What erasing a cell side does: merges the cells on either side of an inner line when they form
 * a rectangle together, reports an outer line, or refuses.
 */
export function eraseTableEdge(table: Table, edge: ErasedEdge, newId: IdFactory): EraseOutcome {
	const grid = tableGrid(table);
	const entry = grid?.rows[edge.row]?.[edge.index];
	if (!grid || !entry) return { kind: 'refused' };
	const origin = mergeOrigin(grid, entry);
	const own = rectOf(grid, origin);
	const along = clampFraction(edge.fraction ?? 0.5);
	let first: GridCell | undefined;
	let second: GridCell | undefined;
	if (edge.side === 'left' || edge.side === 'right') {
		const boundary = edge.side === 'left' ? own.firstCol : own.lastCol + 1;
		if (boundary <= 0 || boundary >= grid.columns) return { kind: 'outer' };
		const height = own.lastRow - own.firstRow + 1;
		const row = own.firstRow + Math.min(height - 1, Math.floor(along * height));
		const a = grid.cover[row]?.[boundary - 1];
		const b = grid.cover[row]?.[boundary];
		first = a && mergeOrigin(grid, a);
		second = b && mergeOrigin(grid, b);
		if (!first || !second || first === second) return { kind: 'refused' };
		const [ra, rb] = [rectOf(grid, first), rectOf(grid, second)];
		if (ra.firstRow !== rb.firstRow || ra.lastRow !== rb.lastRow) return { kind: 'refused' };
	} else {
		const boundary = edge.side === 'top' ? own.firstRow : own.lastRow + 1;
		if (boundary <= 0 || boundary >= grid.rows.length) return { kind: 'outer' };
		const width = own.lastCol - own.firstCol + 1;
		const column = own.firstCol + Math.min(width - 1, Math.floor(along * width));
		const a = grid.cover[boundary - 1]?.[column];
		const b = grid.cover[boundary]?.[column];
		first = a && mergeOrigin(grid, a);
		second = b && mergeOrigin(grid, b);
		if (!first || !second || first === second) return { kind: 'refused' };
		if (first.col !== second.col || first.span !== second.span) return { kind: 'refused' };
	}
	const [ra, rb] = [rectOf(grid, first), rectOf(grid, second)];
	const rect: GridRect = {
		firstRow: Math.min(ra.firstRow, rb.firstRow),
		lastRow: Math.max(ra.lastRow, rb.lastRow),
		firstCol: Math.min(ra.firstCol, rb.firstCol),
		lastCol: Math.max(ra.lastCol, rb.lastCol),
	};
	const merged = mergeCells(table, rect, newId);
	return merged ? { kind: 'merged', table: merged } : { kind: 'refused' };
}
