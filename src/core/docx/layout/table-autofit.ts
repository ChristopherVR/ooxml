// AutoFit table columns (ECMA-376 17.4.53 `w:tblLayout`, 17.4.64 `w:tblW`, 17.4.72 `w:tcW`). The
// rules were fitted to the grids Word saves (docx/__fixtures__/table-autofit/word-grids.json):
//   - every column has a minimum and a maximum content width (table-autofit-measure.ts); a column
//     whose cells have a preferred width takes the largest of them (never below its minimum);
//   - a merged cell spreads its minimum over its columns in proportion to their minimum plus
//     maximum, and its maximum in proportion to their maximum, where it needs more room;
//   - an automatic-width table takes every column at its maximum when they fit, otherwise the
//     columns without a preferred width shrink between their minimum and maximum, and only then
//     the preferred ones; a table with a width grows the columns in proportion to their maximum;
//   - widths are whole twips adding up to the table width.
import type { LayoutBlock, LayoutTable, LayoutTableCell } from './input';
import type { TextMeasurer } from './measure';
import { cellRange, TWIP_PX, type ContentRange } from './table-autofit-measure';
import { distributeColumns, type AutofitColumn } from './table-autofit-distribute';

type Table = Extract<LayoutBlock, { kind: 'table' }>;

export const AUTOFIT_SPAN_PREFERRED_NOTE =
	'AutoFit tables size a merged cell’s preferred width like its content, which has not been measured against Word.';

interface Placed {
	cell: LayoutTableCell;
	start: number;
	span: number;
}

function placedCells(table: Table): Placed[][] {
	return table.rows.map((row) => {
		let start = 0;
		return row.cells.map((cell) => {
			const span = Math.max(1, cell.gridSpan ?? 1);
			const placed = { cell, start, span };
			start += span;
			return placed;
		});
	});
}

/** Spreads `total` over the columns in proportion to `weight`, raising only where it adds room. */
function spread(
	columns: AutofitColumn[],
	total: number,
	weight: (column: AutofitColumn) => number,
	apply: (column: AutofitColumn, share: number) => void,
): void {
	const sum = columns.reduce((acc, column) => acc + weight(column), 0);
	for (const column of columns)
		apply(column, sum > 0 ? (total * weight(column)) / sum : total / columns.length);
}

/** Column minimum, maximum and preferred widths from the cells. */
function measureColumns(
	table: Table,
	count: number,
	measurer: TextMeasurer,
	note: (message: string) => void,
): AutofitColumn[] {
	const columns: AutofitColumn[] = Array.from({ length: count }, () => ({ minPx: 0, maxPx: 0 }));
	const nested = (inner: Table): ContentRange => tableRange(inner, measurer, note);
	const spans: { placed: Placed; range: ContentRange }[] = [];
	for (const row of placedCells(table))
		for (const placed of row) {
			if (placed.start >= count) continue;
			const range = cellRange(placed.cell, measurer, nested);
			if (placed.span > 1) {
				spans.push({ placed, range });
				continue;
			}
			const column = columns[placed.start]!;
			column.minPx = Math.max(column.minPx, range.minPx);
			column.maxPx = Math.max(column.maxPx, range.maxPx);
			const { preferredWidthPx, preferredFraction } = placed.cell;
			if (preferredWidthPx !== undefined)
				column.preferredPx = Math.max(column.preferredPx ?? 0, preferredWidthPx);
			else if (preferredFraction !== undefined)
				column.preferredFraction = Math.max(column.preferredFraction ?? 0, preferredFraction);
		}
	spans.sort((a, b) => a.placed.span - b.placed.span);
	for (const { placed, range } of spans) {
		const covered = columns.slice(placed.start, placed.start + placed.span);
		const preferred = placed.cell.preferredWidthPx;
		if (preferred !== undefined) note(AUTOFIT_SPAN_PREFERRED_NOTE);
		const maxPx = Math.max(range.maxPx, preferred ?? 0);
		if (range.minPx > covered.reduce((sum, column) => sum + column.minPx, 0))
			spread(
				covered,
				range.minPx,
				(column) => column.minPx + column.maxPx,
				(column, share) => {
					column.minPx = Math.max(column.minPx, share);
					column.maxPx = Math.max(column.maxPx, column.minPx);
				},
			);
		if (maxPx > covered.reduce((sum, column) => sum + column.maxPx, 0))
			spread(
				covered,
				maxPx,
				(column) => column.maxPx,
				(column, share) => {
					column.maxPx = Math.max(column.maxPx, share);
				},
			);
	}
	return columns;
}

/** A table's own minimum and maximum widths (a nested table inside a cell). */
function tableRange(table: Table, measurer: TextMeasurer, note: (message: string) => void) {
	if (!table.autofit) {
		const width = table.widthPx ?? 0;
		return { minPx: width, maxPx: width };
	}
	const columns = measureColumns(table, table.autofit.columns, measurer, note);
	const sum = (pick: (column: AutofitColumn) => number) =>
		columns.reduce((acc, column) => acc + pick(column), 0);
	return {
		minPx: sum((column) => column.minPx),
		maxPx: sum((column) => Math.max(column.minPx, column.preferredPx ?? column.maxPx)),
	};
}

const cache = new WeakMap<
	LayoutTable,
	{ availablePx: number; measurer: TextMeasurer; result: LayoutTable }
>();

/**
 * The table with its AutoFit column widths for a text column `availablePx` wide (cell `widthPx`
 * and `xPx`, and the table `widthPx`). Fixed-layout tables are returned unchanged. The result is
 * cached per table, so repeated calls for the same column keep one object.
 */
export function resolveTableAutofit(
	table: LayoutTable,
	availablePx: number,
	measurer: TextMeasurer,
	note: (message: string) => void = () => {},
): LayoutTable {
	const autofit = table.autofit;
	if (!autofit) return table;
	const hit = cache.get(table);
	if (hit && hit.availablePx === availablePx && hit.measurer === measurer) return hit.result;
	const columns = measureColumns(table, autofit.columns, measurer, note);
	const first = table.rows[0]?.cells[0]?.padding;
	const last = table.rows[0]?.cells.at(-1)?.padding;
	const overhang = autofit.legacyMargins ? (first?.left ?? 0) : 0;
	const reference = autofit.legacyMargins
		? availablePx + (first?.left ?? 0) + (last?.right ?? 0)
		: availablePx;
	const widths = distributeColumns(columns, {
		availablePx: availablePx + overhang - Math.max(0, table.indentPx ?? 0),
		...(autofit.widthPx !== undefined
			? { targetPx: autofit.widthPx }
			: autofit.widthFraction !== undefined
				? { targetPx: autofit.widthFraction * reference }
				: {}),
	}).map((twips) => twips * TWIP_PX);
	const offsets = widths.map((_, index) =>
		widths.slice(0, index).reduce((sum, width) => sum + width, 0),
	);
	const placedRows = placedCells(table);
	const result: LayoutTable = {
		...table,
		widthPx: widths.reduce((sum, width) => sum + width, 0),
		rows: table.rows.map((row, rowIndex) => {
			const placed = placedRows[rowIndex]!;
			return {
				...row,
				cells: row.cells.map((cell, cellIndex) => {
					const { start, span } = placed[cellIndex]!;
					if (start >= widths.length) return cell;
					const end = Math.min(widths.length, start + span);
					return {
						...cell,
						xPx: offsets[start]!,
						widthPx: widths.slice(start, end).reduce((sum, width) => sum + width, 0),
					};
				}),
			};
		}),
	};
	cache.set(table, { availablePx, measurer, result });
	return result;
}
