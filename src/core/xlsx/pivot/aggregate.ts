import { rangeContains, type CellAddress } from '../address';
import { type CellValue, type PivotTable, type Worksheet } from '../model';
import { pivotItemLabel, pivotItems, pivotValueKey, validatePivotFields } from './data';
import { hierarchicalPivotGrid } from './hierarchy';
import { pivotHeaders, pivotNumbers, pivotMeasures } from './measures';

export interface PivotGrid {
	values: CellValue[][];
	rowItems: CellValue[];
	columnItems: CellValue[];
	firstDataRow: number;
	rowLabelColumns?: number;
	indents?: number[];
	secondRowItems?: CellValue[];
	rowEntries?: { indices: number[]; type?: 'default' | 'grand' }[];
}

/** Tabular output, including grand totals computed from source records (not averages of cells). */
export function pivotGrid(pivot: PivotTable): PivotGrid {
	validatePivotFields(pivot);
	if (pivot.secondRowField) return hierarchicalPivotGrid(pivot);
	const rowItems = pivotItems(pivot, pivot.rowField);
	const columnItems = pivot.columnField ? pivotItems(pivot, pivot.columnField) : [];
	if (
		(rowItems.length + 3) * ((columnItems.length + 1) * pivotMeasures(pivot).length + 1) >
		1_000_000
	)
		throw new Error('The PivotTable result is too large.');
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
	const label = pivotItemLabel;
	const values = pivotHeaders(
		pivot,
		columnItems,
		[pivot.layout === 'compact' ? 'Row Labels' : pivot.rowField],
		label,
	);
	const firstDataRow = values.length;
	for (const row of rowItems) {
		const records = groups.get(pivotValueKey(row)) ?? [];
		values.push([
			label(row),
			...columnItems.flatMap((col) =>
				pivotNumbers(
					pivot,
					cells.get(JSON.stringify([pivotValueKey(row), pivotValueKey(col)])) ?? [],
				),
			),
			...pivotNumbers(pivot, records),
		]);
	}
	values.push([
		'Grand Total',
		...columnItems.flatMap((col) => pivotNumbers(pivot, columns.get(pivotValueKey(col)) ?? [])),
		...pivotNumbers(pivot, pivot.records),
	]);
	return { values, rowItems, columnItems, firstDataRow };
}

export function pivotTableAt(sheet: Worksheet, at: CellAddress): PivotTable | undefined {
	return sheet.pivotTables?.find((pivot) => rangeContains(pivot.location, at));
}
