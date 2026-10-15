import { rangeContains, type CellAddress } from '../address';
import { type CellValue, type PivotAggregate, type PivotTable, type Worksheet } from '../model';
import { validatePivotFields } from './data';
import { pivotPageProjection, prependPivotPage } from './page-filter';
import type { ColumnEntry } from './columns';
import { nestedPivotGrid } from './nested';

/**
 * One row of `rowItems`: the item path (one index per row level, the measure for the Values
 * level), and `type` for a subtotal or grand total row.
 */
export interface PivotRowEntry {
	indices: number[];
	type?: 'default' | 'grand' | PivotAggregate;
	measure?: number;
}
export interface PivotGrid {
	values: CellValue[][];
	rowItems: CellValue[];
	columnItems: CellValue[];
	firstDataRow: number;
	rowLabelColumns?: number;
	indents?: number[];
	secondRowItems?: CellValue[];
	/** The visible items of every row field, outermost first. */
	axisItems?: CellValue[][];
	/** The visible items of every column field, outermost first. */
	columnAxisItems?: CellValue[][];
	/** The column groups in display order, with their subtotals and the grand total. */
	columnEntries?: ColumnEntry[];
	rowEntries?: PivotRowEntry[];
	firstHeaderRow?: number;
}

/** Pivot output, including grand totals computed from source records (not averages of cells). */
export function pivotGrid(pivot: PivotTable): PivotGrid {
	validatePivotFields(pivot);
	return prependPivotPage(nestedPivotGrid(pivotPageProjection(pivot)), pivot);
}

export function pivotTableAt(sheet: Worksheet, at: CellAddress): PivotTable | undefined {
	return sheet.pivotTables?.find((pivot) => rangeContains(pivot.location, at));
}
