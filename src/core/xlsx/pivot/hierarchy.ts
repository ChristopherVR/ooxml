import type { CellValue, PivotAggregate, PivotTable } from '../model';
import type { PivotGrid } from './aggregate';
import { pivotItemLabel, pivotItems, pivotValueKey } from './data';

type Aggregate = (
	records: readonly CellValue[][],
	field: number,
	type: PivotAggregate,
) => CellValue;
const label = pivotItemLabel;

/** Native tabular two-field layout: actual tuples, bottom automatic subtotals and grand totals. */
export function hierarchicalPivotGrid(pivot: PivotTable, aggregate: Aggregate): PivotGrid {
	const rowItems = pivotItems(pivot, pivot.rowField);
	const secondRowItems = pivotItems(pivot, pivot.secondRowField!);
	const columnItems = pivot.columnField ? pivotItems(pivot, pivot.columnField) : [];
	const outerField = pivot.fields.indexOf(pivot.rowField);
	const innerField = pivot.fields.indexOf(pivot.secondRowField!);
	const columnField = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const valueField = pivot.fields.indexOf(pivot.valueField);
	const groups = new Map<string, CellValue[][]>();
	const details = new Map<string, CellValue[][]>();
	const children = new Map<string, Set<string>>();
	const innerIndices = new Map(secondRowItems.map((item, index) => [pivotValueKey(item), index]));
	const buckets = new Map<string, CellValue[][]>();
	const collect = (map: Map<string, CellValue[][]>, key: string, record: CellValue[]) => {
		const list = map.get(key);
		if (list) list.push(record);
		else map.set(key, [record]);
	};
	for (const record of pivot.records) {
		const outer = pivotValueKey(record[outerField] ?? null);
		const inner = pivotValueKey(record[innerField] ?? null);
		collect(groups, outer, record);
		collect(details, JSON.stringify([outer, inner]), record);
		if (!children.has(outer)) children.set(outer, new Set());
		children.get(outer)!.add(inner);
		if (pivot.columnField) {
			const column = pivotValueKey(record[columnField] ?? null);
			for (const key of [[], [outer], [outer, inner]])
				collect(buckets, JSON.stringify([...key, column]), record);
		}
	}
	const rows = details.size + (pivot.subtotals !== false ? rowItems.length : 0) + 3;
	if (rows * (columnItems.length + 3) > 1_000_000)
		throw new Error('The PivotTable result is too large.');
	const values: CellValue[][] = pivot.columnField
		? [
				[
					pivot.valueCaption,
					null,
					pivot.columnField,
					...columnItems.slice(1).map(() => null),
					null,
				],
				[pivot.rowField, pivot.secondRowField!, ...columnItems.map(label), 'Grand Total'],
			]
		: [[pivot.rowField, pivot.secondRowField!, pivot.valueCaption]];
	const rowEntries: NonNullable<PivotGrid['rowEntries']> = [];
	const numbers = (keys: string[], records: CellValue[][]): CellValue[] => [
		...columnItems.map((column) =>
			aggregate(
				buckets.get(JSON.stringify([...keys, pivotValueKey(column)])) ?? [],
				valueField,
				pivot.aggregate,
			),
		),
		aggregate(records, valueField, pivot.aggregate),
	];
	rowItems.forEach((outer, outerIndex) => {
		const key = pivotValueKey(outer);
		let first = true;
		[...children.get(key)!]
			.sort((a, b) => innerIndices.get(a)! - innerIndices.get(b)!)
			.forEach((innerKey) => {
				const innerIndex = innerIndices.get(innerKey)!;
				const inner = secondRowItems[innerIndex]!;
				const keys = [key, innerKey];
				const records = details.get(JSON.stringify(keys));
				if (!records) return;
				values.push([
					first || pivot.repeatLabels ? label(outer) : null,
					label(inner),
					...numbers(keys, records),
				]);
				rowEntries.push({ indices: [outerIndex, innerIndex] });
				first = false;
			});
		if (pivot.subtotals !== false) {
			const caption = String(label(outer));
			values.push([`${caption} Total`, null, ...numbers([key], groups.get(key) ?? [])]);
			rowEntries.push({ indices: [outerIndex], type: 'default' });
		}
	});
	values.push(['Grand Total', null, ...numbers([], pivot.records)]);
	rowEntries.push({ indices: [0], type: 'grand' });
	return {
		values,
		rowItems,
		secondRowItems,
		columnItems,
		rowEntries,
		firstDataRow: pivot.columnField ? 2 : 1,
	};
}
