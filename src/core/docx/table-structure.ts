// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Row and column edits on tables with merged cells (`gridSpan`, `vMerge`). Each edit takes the
// table model and returns a new one, or undefined when the table cannot take the edit (not
// rectangular on its grid, or the edit would leave no table). The results follow what Word
// produces for the same edit; src/core/docx/__fixtures__/word-table-structure.json records them
// and table-structure.test.ts checks every one.
import type { Table } from './model';
import {
	copiedRowProperties,
	copyTable,
	emptyCell,
	paddedRowProperties,
	storeRowProperties,
	type IdFactory,
} from './table-structure-cells';
import { continuationsOf, tableGrid } from './table-grid';
import type { Twips } from './units';

/** Removes a vertical merge that no longer spans two rows. */
export function normalizeMerges(table: Table): void {
	table.rows.forEach((row, r) =>
		row.forEach((cell, index) => {
			if (cell.verticalMerge !== 'restart') return;
			const grid = tableGrid(table);
			const entry = grid?.rows[r]?.[index];
			if (grid && entry && !continuationsOf(grid, entry).length) delete cell.verticalMerge;
		}),
	);
}

/**
 * Inserts a row above or below row `row`, shaped like it. A cell inside a vertical merge that
 * continues through the new row is part of that merge; any other new cell stands alone.
 */
export function insertRow(
	table: Table,
	row: number,
	where: 'above' | 'below',
	newId: IdFactory,
): Table | undefined {
	const grid = tableGrid(table);
	const reference = grid?.rows[row];
	if (!grid || !reference) return undefined;
	const at = where === 'above' ? row : row + 1;
	const next = copyTable(table);
	const added = reference.map((entry) => {
		const cell = emptyCell(newId, entry.cell, entry.span, entry.cell.widthTwips);
		const following = grid.cover[at]?.[entry.col];
		if (following?.cell.verticalMerge === 'continue' && following.span === entry.span)
			cell.verticalMerge = 'continue';
		return cell;
	});
	const properties = paddedRowProperties(table);
	next.rows.splice(at, 0, added);
	properties.splice(at, 0, copiedRowProperties(table, row));
	storeRowProperties(next, properties);
	return next;
}

/**
 * Deletes row `row`. A vertical merge that started in it starts in the next row instead; one
 * left with a single row becomes an ordinary cell.
 */
export function deleteRow(table: Table, row: number): Table | undefined {
	const grid = tableGrid(table);
	const cells = grid?.rows[row];
	if (!grid || !cells || table.rows.length < 2) return undefined;
	const next = copyTable(table);
	for (const entry of cells) {
		if (entry.cell.verticalMerge !== 'restart') continue;
		const [first] = continuationsOf(grid, entry);
		const promoted = first && next.rows[first.row]?.[first.index];
		if (promoted) promoted.verticalMerge = 'restart';
	}
	const properties = paddedRowProperties(table);
	next.rows.splice(row, 1);
	properties.splice(row, 1);
	storeRowProperties(next, properties);
	normalizeMerges(next);
	return next;
}

const widthOf = (table: Table, col: number): Twips | undefined => table.grid?.[col];

/**
 * Inserts a column to the left or right of grid column `col`, as Word does: every row gets a new
 * one-column cell beside the cell that holds `col`, and the new grid column takes `col`'s width
 * (the table grows; no other width changes). Beside a merged cell that reaches past `col`, a row
 * gets its new cell after the merged cell, so the new column is not straight there. Word's result
 * for a column inserted to the left of, or through, a merged cell is not a plain grid (it adds
 * `gridAfter` rows), so that edit, and any edit that would leave a vertical merge unaligned, is refused.
 */
export function insertColumn(
	table: Table,
	col: number,
	where: 'left' | 'right',
	newId: IdFactory,
): Table | undefined {
	const grid = tableGrid(table);
	if (!grid || col < 0 || col >= grid.columns) return undefined;
	const next = copyTable(table);
	const width = widthOf(table, col);
	for (let r = 0; r < grid.rows.length; r++) {
		const anchor = grid.cover[r]?.[col];
		const target = next.rows[r];
		if (!anchor || !target) return undefined;
		if (where === 'left' && (anchor.col !== col || anchor.span > 1)) return undefined;
		const created = emptyCell(
			newId,
			anchor.cell,
			1,
			anchor.cell.widthTwips !== undefined ? width : undefined,
		);
		delete created.verticalMerge;
		target.splice(anchor.index + (where === 'right' ? 1 : 0), 0, created);
	}
	if (next.grid && width !== undefined)
		next.grid.splice(col + (where === 'right' ? 1 : 0), 0, width);
	return tableGrid(next) ? next : undefined;
}

/**
 * Deletes grid columns `first` to `last`, with every cell in them. Word deletes a merged cell that
 * is only partly in the columns whole and leaves the row short (`gridAfter`), which a plain grid
 * cannot hold, so that edit is refused, as is deleting every column.
 */
export function deleteColumns(table: Table, first: number, last: number): Table | undefined {
	const grid = tableGrid(table);
	if (!grid || first < 0 || last >= grid.columns || first > last) return undefined;
	if (last - first + 1 >= grid.columns) return undefined;
	const next = copyTable(table);
	for (let r = 0; r < grid.rows.length; r++) {
		const cells = grid.rows[r] ?? [];
		const inside = cells.filter(
			(entry) => entry.col <= last && entry.col + entry.span - 1 >= first,
		);
		if (inside.some((entry) => entry.col < first || entry.col + entry.span - 1 > last))
			return undefined;
		const kept = new Set(inside.map((entry) => entry.index));
		const target = next.rows[r];
		if (!target) return undefined;
		next.rows[r] = target.filter((_, index) => !kept.has(index));
	}
	if (next.grid) next.grid.splice(first, last - first + 1);
	normalizeMerges(next);
	return tableGrid(next) ? next : undefined;
}
