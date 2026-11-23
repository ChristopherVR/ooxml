// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Pure editing operations for table-level properties (`w:tblW`, `w:jc`, `w:tblInd`,
// `w:tblBorders`, `w:tblStyle`, `w:tblLook`, and `w:tblGrid` / `w:gridCol` widths).
import type { Table } from './model';
import type { TableBorders, TableBorderSide, TableLook } from './table-model';
import { tableColumnCount } from './table-defaults';
import { spanOf, tableGrid } from './table-grid';
import { signedTwips, twips, type Twips } from './units';

/** Preferred table width specification (`w:tblW`). For `pct`, `value` is a positive fraction (e.g. `0.5` for 50%). */
export type TableWidthSpec =
	| { type: 'auto' }
	| { type: 'dxa'; value: number }
	| { type: 'pct'; value: number };

/** Partial update for table borders; setting a side to `undefined` clears that side. */
export type TableBordersPatch = Partial<Record<keyof TableBorders, TableBorderSide | undefined>>;

/** Partial update for table look flags; setting a flag to `undefined` clears that flag. */
export type TableLookPatch = Partial<Record<keyof TableLook, boolean | undefined>>;

/** Combined patch for updating one or more table-level properties at once. */
export interface TablePropertiesPatch {
	width?: TableWidthSpec;
	alignment?: Table['alignment'] | undefined;
	indentTwips?: number | undefined;
	borders?: TableBordersPatch | undefined;
	style?: string | undefined;
	look?: TableLookPatch | undefined;
	gridWidths?: readonly number[];
}

const BORDER_SIDES: readonly (keyof TableBorders)[] = [
	'top',
	'bottom',
	'left',
	'right',
	'insideH',
	'insideV',
];

const LOOK_FLAGS: readonly (keyof TableLook)[] = [
	'firstRow',
	'lastRow',
	'firstColumn',
	'lastColumn',
	'noHBand',
	'noVBand',
];

/** Updates the table's preferred width (`w:tblW`: `dxa`, `pct`, or `auto`). */
export function setTableWidth(table: Table, width: TableWidthSpec): Table {
	const next = structuredClone(table);
	if (width.type === 'auto') {
		delete next.widthTwips;
		delete next.widthPercent;
		return next;
	}
	if (width.type === 'dxa') {
		if (!Number.isFinite(width.value) || width.value < 0)
			throw new RangeError('Table width in twips must be a non-negative finite number.');
		next.widthTwips = twips(Math.round(width.value));
		delete next.widthPercent;
		return next;
	}
	if (!Number.isFinite(width.value) || width.value <= 0)
		throw new RangeError('Table percentage width must be a positive finite fraction.');
	next.widthPercent = width.value;
	delete next.widthTwips;
	return next;
}

/** Updates the table's horizontal alignment (`w:jc`: `left`, `center`, `right`, or `undefined` to clear). */
export function setTableAlignment(table: Table, alignment: Table['alignment'] | undefined): Table {
	const next = structuredClone(table);
	if (alignment === undefined) {
		delete next.alignment;
		delete next.justification;
		return next;
	}
	next.alignment = alignment;
	next.justification = alignment;
	return next;
}

/** Updates the table's leading-edge indentation in twips (`w:tblInd`), or clears it when `undefined`. */
export function setTableIndent(table: Table, indentTwips: number | undefined): Table {
	const next = structuredClone(table);
	if (indentTwips === undefined) {
		delete next.indentTwips;
		return next;
	}
	if (!Number.isFinite(indentTwips))
		throw new RangeError('Table indent in twips must be a finite number.');
	next.indentTwips = signedTwips(Math.round(indentTwips));
	return next;
}

/**
 * Merges border side updates into `w:tblBorders`. Pass `undefined` to clear all table borders, or
 * set an individual side to `undefined` to clear only that side.
 */
export function setTableBorders(table: Table, borders: TableBordersPatch | undefined): Table {
	const next = structuredClone(table);
	if (borders === undefined) {
		delete next.borders;
		return next;
	}
	const merged: TableBorders = { ...next.borders };
	for (const side of BORDER_SIDES) {
		if (!(side in borders)) continue;
		const value = borders[side];
		if (value === undefined) delete merged[side];
		else merged[side] = structuredClone(value);
	}
	if (Object.keys(merged).length) next.borders = merged;
	else delete next.borders;
	return next;
}

/** Updates the table's style reference (`w:tblStyle`), or clears it when `undefined` or blank. */
export function setTableStyle(table: Table, style: string | undefined): Table {
	const next = structuredClone(table);
	const trimmed = style?.trim();
	if (!trimmed) delete next.style;
	else next.style = trimmed;
	return next;
}

/**
 * Merges conditional formatting flags into `w:tblLook`. Pass `undefined` to clear all look flags,
 * or set an individual flag to `undefined` to clear only that flag.
 */
export function setTableLook(table: Table, look: TableLookPatch | undefined): Table {
	const next = structuredClone(table);
	if (look === undefined) {
		delete next.look;
		return next;
	}
	const merged: TableLook = { ...next.look };
	for (const flag of LOOK_FLAGS) {
		if (!(flag in look)) continue;
		const value = look[flag];
		if (value === undefined) delete merged[flag];
		else merged[flag] = value;
	}
	if (Object.keys(merged).length) next.look = merged;
	else delete next.look;
	return next;
}

/**
 * Updates `w:tblGrid` column widths in twips and keeps every cell's `widthTwips` consistent with
 * the sum of the grid columns it spans (`gridSpan`).
 */
export function setTableGridWidths(table: Table, widths: readonly number[]): Table {
	const expectedColumns = tableGrid(table)?.columns ?? tableColumnCount(table);
	if (!widths.length || widths.length !== expectedColumns)
		throw new RangeError(
			`Expected ${expectedColumns} grid column width(s), received ${widths.length}.`,
		);
	const normalized: Twips[] = widths.map((value) => {
		if (!Number.isFinite(value) || value <= 0)
			throw new RangeError('Each grid column width must be a positive finite number of twips.');
		return twips(Math.round(value));
	});
	const next = structuredClone(table);
	next.grid = normalized;
	const grid = tableGrid(next);
	if (grid) {
		for (const row of grid.rows)
			for (const entry of row) {
				const sum = normalized
					.slice(entry.col, entry.col + entry.span)
					.reduce((total, w) => total + w, 0);
				entry.cell.widthTwips = twips(sum);
				delete entry.cell.widthPercent;
			}
	} else {
		for (const row of next.rows) {
			let col = 0;
			for (const cell of row) {
				const span = spanOf(cell);
				const sum = normalized.slice(col, col + span).reduce((total, w) => total + w, 0);
				if (sum > 0) {
					cell.widthTwips = twips(sum);
					delete cell.widthPercent;
				}
				col += span;
			}
		}
	}
	return next;
}

/** Applies a batch of table-level property updates, returning a new `Table` without mutating `table`. */
export function updateTableProperties(table: Table, patch: TablePropertiesPatch): Table {
	let next = structuredClone(table);
	if (patch.width !== undefined) next = setTableWidth(next, patch.width);
	if ('alignment' in patch) next = setTableAlignment(next, patch.alignment);
	if ('indentTwips' in patch) next = setTableIndent(next, patch.indentTwips);
	if ('borders' in patch) next = setTableBorders(next, patch.borders);
	if ('style' in patch) next = setTableStyle(next, patch.style);
	if ('look' in patch) next = setTableLook(next, patch.look);
	if (patch.gridWidths !== undefined) next = setTableGridWidths(next, patch.gridWidths);
	return next;
}
