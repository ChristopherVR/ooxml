// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The grid view of a table: where each physical cell sits on the grid columns, with `gridSpan`
// widths and `vMerge` runs resolved. Structure edits (table-structure.ts) and the editor's
// row-span calculation both read tables through it.
import type { Table, TableCell } from './model';

/** One physical cell and the grid area it covers. */
export interface GridCell {
	cell: TableCell;
	/** Row index in `Table.rows`. */
	row: number;
	/** Position in that row's cell array. */
	index: number;
	/** First grid column covered. */
	col: number;
	/** Grid columns covered (`gridSpan`, at least 1). */
	span: number;
}

export interface TableGrid {
	/** Number of grid columns. */
	columns: number;
	/** The cells of each row in order. */
	rows: GridCell[][];
	/** `cover[row][col]`: the physical cell occupying that grid position. */
	cover: (GridCell | undefined)[][];
}

export const spanOf = (cell: TableCell): number => Math.max(1, cell.gridSpan ?? 1);

/**
 * Lays the table out on its grid, or returns undefined when it is not a consistent rectangle of
 * grid positions (a row shorter or longer than the grid, a row offset the model cannot express,
 * or a vertical merge continuation with no matching cell above it).
 */
export function tableGrid(table: Table): TableGrid | undefined {
	const columns = table.grid?.length ?? sumSpans(table.rows[0]);
	if (!columns || !table.rows.length) return undefined;
	const rows: GridCell[][] = [];
	const cover: (GridCell | undefined)[][] = [];
	for (let row = 0; row < table.rows.length; row++) {
		const cells = table.rows[row] ?? [];
		if (!cells.length || sumSpans(cells) !== columns) return undefined;
		const placed: GridCell[] = [];
		const line: (GridCell | undefined)[] = new Array(columns);
		let col = 0;
		cells.forEach((cell, index) => {
			const entry: GridCell = { cell, row, index, col, span: spanOf(cell) };
			for (let k = 0; k < entry.span; k++) line[col + k] = entry;
			col += entry.span;
			placed.push(entry);
		});
		rows.push(placed);
		cover.push(line);
	}
	const grid: TableGrid = { columns, rows, cover };
	for (const entry of rows.flat())
		if (entry.cell.verticalMerge === 'continue') {
			const above = entry.row > 0 ? cover[entry.row - 1]?.[entry.col] : undefined;
			if (!above || above.col !== entry.col || above.span !== entry.span) return undefined;
			if (!above.cell.verticalMerge) return undefined;
		}
	return grid;
}

function sumSpans(cells: TableCell[] | undefined): number {
	return (cells ?? []).reduce((sum, cell) => sum + spanOf(cell), 0);
}

/** The cell that starts the vertical merge `entry` belongs to (itself when it is not merged). */
export function mergeOrigin(grid: TableGrid, entry: GridCell): GridCell {
	let current = entry;
	while (current.cell.verticalMerge === 'continue' && current.row > 0) {
		const above = grid.cover[current.row - 1]?.[current.col];
		if (!above) break;
		current = above;
	}
	return current;
}

/** Rows a cell covers: 1 for an unmerged cell, more for the start of a vertical merge. */
export function rowSpanOf(grid: TableGrid, entry: GridCell): number {
	if (entry.cell.verticalMerge !== 'restart') return 1;
	let height = 1;
	while (grid.cover[entry.row + height]?.[entry.col]?.cell.verticalMerge === 'continue') height++;
	return height;
}

/** The continuation cells below a merge start, top to bottom. */
export function continuationsOf(grid: TableGrid, entry: GridCell): GridCell[] {
	const found: GridCell[] = [];
	for (let r = entry.row + 1; ; r++) {
		const below = grid.cover[r]?.[entry.col];
		if (below?.cell.verticalMerge !== 'continue') return found;
		found.push(below);
	}
}
