import type { Table } from './model';
import { newTableGridWidths } from './table-defaults';
import { setTableGridWidths } from './table-properties-edit';
import { twips } from './units';

export type TableAutoFit = 'contents' | 'window' | 'fixed';

/** AutoFit uses the same preferred-width inputs as the print layout engine. */
export function autoFitTable(table: Table, mode: TableAutoFit): Table {
	const next = structuredClone(table);
	next.layout = mode === 'fixed' ? 'fixed' : 'autofit';
	if (mode !== 'fixed') {
		delete next.widthTwips;
		delete next.widthPercent;
		if (mode === 'window') next.widthPercent = 1;
		for (const row of next.rows)
			for (const cell of row) {
				delete cell.widthTwips;
				delete cell.widthPercent;
			}
	}
	return next;
}

/** Sets selected grid columns to a width, or divides their existing total evenly. */
export function resizeTableColumns(
	table: Table,
	first: number,
	last: number,
	contentWidthTwips: number,
	width?: number,
): Table {
	const fallback = table.widthTwips ?? contentWidthTwips * (table.widthPercent ?? 1);
	const grid = [...newTableGridWidths(table, fallback)];
	const selected = grid.slice(first, last + 1);
	if (first < 0 || last >= grid.length || !selected.length) return table;
	const total = selected.reduce((sum, value) => sum + value, 0);
	const even = Math.floor(total / selected.length);
	for (let col = first; col <= last; col++)
		grid[col] = width ?? Math.max(1, even + Number(col - first < total % selected.length));
	const next = setTableGridWidths(table, grid, { first, last });
	next.widthTwips = twips(grid.reduce((sum, value) => sum + value, 0));
	delete next.widthPercent;
	next.layout = 'fixed';
	return next;
}

export function resizeTableRows(
	table: Table,
	first: number,
	last: number,
	measured: readonly number[],
	height?: number,
): Table {
	const next = { ...table };
	const selected = measured.slice(first, last + 1);
	if (first < 0 || last >= table.rows.length || !selected.length) return table;
	const total = selected.reduce((sum, value) => sum + value, 0);
	const even = Math.floor(total / selected.length);
	next.rowProperties = table.rows.map((_, index) => {
		const previous = table.rowProperties?.[index] ?? {};
		if (index < first || index > last) return previous;
		return {
			...previous,
			heightTwips: twips(height ?? even + Number(index - first < total % selected.length)),
			heightRule: 'atLeast',
		};
	});
	return next;
}

/** Header rows must start at the first row or continue the existing header group. */
export function canRepeatTableRows(table: Table, first: number, last: number): boolean {
	return (
		first >= 0 &&
		first <= last &&
		last < table.rows.length &&
		Array.from({ length: first }, (_, index) => table.rowProperties?.[index]?.header).every(Boolean)
	);
}

export function repeatTableRows(table: Table, first: number, last: number): Table {
	if (!canRepeatTableRows(table, first, last)) return table;
	const enable = Array.from({ length: last - first + 1 }, (_, i) => i + first).some(
		(index) => !table.rowProperties?.[index]?.header,
	);
	return {
		...table,
		rowProperties: table.rows.map((_, index) => {
			const previous = table.rowProperties?.[index] ?? {};
			return index >= first && (index <= last || (!enable && previous.header))
				? { ...previous, header: enable }
				: previous;
		}),
	};
}
