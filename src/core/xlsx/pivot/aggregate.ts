import { rangeContains, type CellAddress } from '../address';
import {
	isCellError,
	type CellValue,
	type PivotAggregate,
	type PivotTable,
	type Worksheet,
} from '../model';
import { pivotItemLabel, pivotItems, pivotValueKey, validatePivotFields } from './data';

function aggregate(
	records: readonly CellValue[][],
	field: number,
	type: PivotAggregate,
): CellValue {
	if (!records.length) return null;
	const values = records.map((record) => record[field] ?? null);
	if (type === 'count') return values.filter((value) => value !== null && value !== '').length;
	const error = values.find(isCellError);
	if (error) return error;
	const numbers = values.filter((value): value is number => typeof value === 'number');
	if (!numbers.length) return type === 'average' ? { error: '#DIV/0!' } : 0;
	if (type === 'min') return numbers.reduce((a, b) => Math.min(a, b));
	if (type === 'max') return numbers.reduce((a, b) => Math.max(a, b));
	const sum = numbers.reduce((total, value) => total + value, 0);
	return type === 'average' ? sum / numbers.length : sum;
}

export interface PivotGrid {
	values: CellValue[][];
	rowItems: CellValue[];
	columnItems: CellValue[];
	firstDataRow: number;
}

/** Tabular output, including grand totals computed from source records (not averages of cells). */
export function pivotGrid(pivot: PivotTable): PivotGrid {
	validatePivotFields(pivot);
	const rowItems = pivotItems(pivot, pivot.rowField);
	const columnItems = pivot.columnField ? pivotItems(pivot, pivot.columnField) : [];
	if ((rowItems.length + 3) * (columnItems.length + 2) > 1_000_000)
		throw new Error('The PivotTable result is too large.');
	const rowField = pivot.fields.indexOf(pivot.rowField);
	const colField = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const valueField = pivot.fields.indexOf(pivot.valueField);
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
	const values: CellValue[][] = pivot.columnField
		? [
				[pivot.valueCaption, pivot.columnField, ...columnItems.slice(1).map(() => null), null],
				[pivot.rowField, ...columnItems.map(label), 'Grand Total'],
			]
		: [[pivot.rowField, pivot.valueCaption]];
	for (const row of rowItems) {
		const records = groups.get(pivotValueKey(row)) ?? [];
		values.push([
			label(row),
			...columnItems.map((col) =>
				aggregate(
					cells.get(JSON.stringify([pivotValueKey(row), pivotValueKey(col)])) ?? [],
					valueField,
					pivot.aggregate,
				),
			),
			aggregate(records, valueField, pivot.aggregate),
		]);
	}
	values.push([
		'Grand Total',
		...columnItems.map((col) =>
			aggregate(columns.get(pivotValueKey(col)) ?? [], valueField, pivot.aggregate),
		),
		aggregate(pivot.records, valueField, pivot.aggregate),
	]);
	return { values, rowItems, columnItems, firstDataRow: pivot.columnField ? 2 : 1 };
}

export function pivotTableAt(sheet: Worksheet, at: CellAddress): PivotTable | undefined {
	return sheet.pivotTables?.find((pivot) => rangeContains(pivot.location, at));
}
