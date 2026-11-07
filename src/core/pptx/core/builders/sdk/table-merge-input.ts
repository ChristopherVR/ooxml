/**
 * The grid of a table made in code: its column widths as fractions, and, with
 * `TableInput.fillMergedCells`, its real cells placed on the grid.
 *
 * OOXML keeps an `a:tc` for every grid slot, and a slot a merge covers is an
 * empty cell marked `hMerge` (to the right of the merge's first cell) and / or
 * `vMerge` (below it). With `fillMergedCells` the caller lists only the real
 * cells and the covered slots are added here.
 *
 * @module sdk/table-merge-input
 */

import type { TableCellInput, TableInput } from './types';

/** One grid slot: a real cell, or a slot a merge covers. */
export type TableSlot =
	| { cell: TableCellInput }
	| { cell?: undefined; hMerge: boolean; vMerge: boolean };

/** The table's real cells placed on its grid, every row `columnCount` slots long. */
export interface TableGrid {
	columnCount: number;
	rows: TableSlot[][];
}

function span(value: number | undefined, name: string): number {
	if (value === undefined) {
		return 1;
	}
	if (!Number.isInteger(value) || value < 1) {
		throw new RangeError(`TableCellInput.${name} must be a whole number from 1, got ${value}`);
	}
	return value;
}

/**
 * Place each row's cells left to right in the slots that earlier merges
 * leave free; a cell spanning `gridSpan` x `rowSpan` covers the slots to its
 * right and below. The column count is `columnWidths.length` when given,
 * else the widest row; a short row is padded with empty cells.
 *
 * @throws RangeError for a span that is not a whole number from 1, a
 *   `rowSpan` that reaches past the last row, a cell that lands on a slot an
 *   earlier merge covers, or a row wider than `columnWidths`.
 */
export function placeTableCells(input: TableInput): TableGrid {
	const rowCount = input.rows.length;
	const slots: Array<Array<TableSlot | undefined>> = input.rows.map(() => []);
	input.rows.forEach((row, ri) => {
		const rowSlots = slots[ri] ?? [];
		let ci = 0;
		for (const cell of row.cells) {
			while (rowSlots[ci]) {
				ci += 1;
			}
			const colSpan = span(cell.gridSpan, 'gridSpan');
			const rowSpan = span(cell.rowSpan, 'rowSpan');
			if (ri + rowSpan > rowCount) {
				throw new RangeError(
					`A cell in row ${ri} spans ${rowSpan} rows, past the table's ${rowCount} rows`,
				);
			}
			for (let dr = 0; dr < rowSpan; dr++) {
				for (let dc = 0; dc < colSpan; dc++) {
					const target = slots[ri + dr] ?? [];
					if (target[ci + dc]) {
						throw new RangeError(
							`A cell in row ${ri} overlaps the merge covering row ${ri + dr}, column ${ci + dc}`,
						);
					}
					target[ci + dc] = dr === 0 && dc === 0 ? { cell } : { hMerge: dc > 0, vMerge: dr > 0 };
				}
			}
			ci += colSpan;
		}
	});
	const widest = Math.max(...slots.map((row) => row.length), 1);
	const declared = input.columnWidths?.length;
	if (declared !== undefined && widest > declared) {
		throw new RangeError(`A row takes ${widest} columns, but columnWidths gives ${declared}`);
	}
	const columnCount = declared ?? widest;
	return {
		columnCount,
		rows: slots.map((row) =>
			Array.from({ length: columnCount }, (_, ci) => row[ci] ?? { cell: { text: '' } }),
		),
	};
}

/**
 * Column widths as fractions of the table's width: `columnWidths` scaled to
 * add up to 1, or equal columns without it.
 *
 * @throws RangeError when `columnWidths` does not give one finite width of 0
 *   or more per column, or its widths add up to 0.
 */
export function columnFractions(columnWidths: number[] | undefined, columnCount: number): number[] {
	if (!columnWidths) {
		return Array.from({ length: columnCount }, () => 1 / columnCount);
	}
	if (columnWidths.length !== columnCount) {
		throw new RangeError(
			`columnWidths gives ${columnWidths.length} widths for a table of ${columnCount} columns`,
		);
	}
	if (columnWidths.some((w) => !Number.isFinite(w) || w < 0)) {
		throw new RangeError('columnWidths must be finite widths of 0 or more');
	}
	const sum = columnWidths.reduce((a, b) => a + b, 0);
	if (sum <= 0) {
		throw new RangeError('columnWidths must not all be 0');
	}
	return columnWidths.map((w) => w / sum);
}
