import { rangeContains, type CellAddress } from '../address';
import { type CellValue, type PivotAggregate, type PivotTable, type Worksheet } from '../model';
import { pivotItemLabel, pivotItems, pivotValueKey, validatePivotFields } from './data';
import { pivotPageProjection, prependPivotPage } from './page-filter';
import { hierarchicalPivotGrid } from './hierarchy';
import type { ColumnEntry } from './columns';
import { nestedPivotGrid } from './nested';
import { pivotHeaderDimensions, pivotHeaders, pivotNumbers, pivotMeasures } from './measures';

/** One row of `rowItems`: the item path, and `type` for a subtotal or grand total row. */
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
	/** The visible items of every row field, outermost first (Values on columns). */
	axisItems?: CellValue[][];
	/** The visible items of every column field, outermost first (Values on columns). */
	columnAxisItems?: CellValue[][];
	/** The column groups in display order, with their subtotals and the grand total. */
	columnEntries?: ColumnEntry[];
	rowEntries?: PivotRowEntry[];
	firstHeaderRow?: number;
}

/** Pivot output, including grand totals computed from source records (not averages of cells). */
export function pivotGrid(pivot: PivotTable): PivotGrid {
	validatePivotFields(pivot);
	return prependPivotPage(pivotDataGrid(pivotPageProjection(pivot)), pivot);
}
function pivotDataGrid(pivot: PivotTable): PivotGrid {
	if (pivot.valuesAxis !== 'rows') return nestedPivotGrid(pivot);
	if (pivot.secondRowField) return hierarchicalPivotGrid(pivot);
	// Values on rows with one row field: every measure is a row under each item.
	const rowItems = pivotItems(pivot, pivot.rowField);
	const columnItems = pivot.columnField ? pivotItems(pivot, pivot.columnField) : [];
	const compact = pivot.layout === 'compact';
	const headings = pivot.layout !== undefined && pivot.layout !== 'tabular';
	const dimensions = pivotHeaderDimensions(pivot, columnItems.length, compact ? 1 : 2);
	const dataRows =
		(rowItems.length + 1) * (1 + (pivot.additionalValues?.length ?? 0)) +
		(headings ? rowItems.length : 0);
	if ((dataRows + dimensions.rows) * dimensions.width > 1_000_000)
		throw new Error('The PivotTable result is too large.');
	const label = pivotItemLabel;
	const values = pivotHeaders(pivot, columnItems, [compact ? 'Row Labels' : pivot.rowField], label);
	const rowField = pivot.fields.indexOf(pivot.rowField);
	const colField = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const groups = new Map<string, CellValue[][]>();
	const cells = new Map<string, CellValue[][]>();
	const columns = new Map<string, CellValue[][]>();
	const collect = (map: Map<string, CellValue[][]>, key: string, record: CellValue[]) => {
		const found = map.get(key);
		if (found) found.push(record);
		else map.set(key, [record]);
	};
	for (const record of pivot.records) {
		const key = pivotValueKey(record[rowField] ?? null);
		collect(groups, key, record);
		if (pivot.columnField) {
			const col = pivotValueKey(record[colField] ?? null);
			collect(columns, col, record);
			collect(cells, JSON.stringify([key, col]), record);
		}
	}
	const firstDataRow = values.length;
	const measures = pivotMeasures(pivot);
	const rowEntries: PivotRowEntry[] = [];
	const indents = values.map(() => 0);
	const append = (caption: CellValue, bands: CellValue[][], index: number, grand = false) => {
		measures.forEach((measure, m) => {
			values.push([
				...(compact
					? [grand ? `Total ${measure.caption}` : measure.caption]
					: [
							grand
								? `Total ${measure.caption}`
								: pivot.repeatLabels || (!headings && m === 0)
									? caption
									: null,
							grand ? null : measure.caption,
						]),
				...bands.map((band) => band[m] ?? null),
			]);
			indents.push(compact && !grand ? 1 : 0);
			rowEntries.push({
				indices: grand ? [0] : [index, m],
				measure: m,
				...(grand ? { type: 'grand' as const } : {}),
			});
		});
	};
	rowItems.forEach((item, index) => {
		const key = pivotValueKey(item);
		if (headings) {
			values.push([label(item), ...Array<CellValue>(dimensions.width - 1).fill(null)]);
			rowEntries.push({ indices: [index] });
			indents.push(0);
		}
		append(
			label(item),
			[
				...columnItems.map((col) =>
					pivotNumbers(pivot, cells.get(JSON.stringify([key, pivotValueKey(col)])) ?? []),
				),
				pivotNumbers(pivot, groups.get(key) ?? []),
			],
			index,
		);
	});
	append(
		'Grand Total',
		[
			...columnItems.map((col) => pivotNumbers(pivot, columns.get(pivotValueKey(col)) ?? [])),
			pivotNumbers(pivot, pivot.records),
		],
		0,
		true,
	);
	return {
		values,
		rowItems,
		columnItems,
		firstDataRow,
		rowLabelColumns: compact ? 1 : 2,
		indents,
		firstHeaderRow: 1,
		rowEntries,
	};
}

export function pivotTableAt(sheet: Worksheet, at: CellAddress): PivotTable | undefined {
	return sheet.pivotTables?.find((pivot) => rangeContains(pivot.location, at));
}
