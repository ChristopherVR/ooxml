import type { CellValue, PivotTable } from '../model';
import type { PivotGrid } from './aggregate';
import { pivotItemLabel, pivotItems, pivotValueKey } from './data';

import { pivotHeaders, pivotNumbers, pivotMeasures } from './measures';
const label = pivotItemLabel;

/** Two-field report layouts share the same indexed tuples and aggregate buckets. */
export function hierarchicalPivotGrid(pivot: PivotTable): PivotGrid {
	const rowItems = pivotItems(pivot, pivot.rowField);
	const secondRowItems = pivotItems(pivot, pivot.secondRowField!);
	const columnItems = pivot.columnField ? pivotItems(pivot, pivot.columnField) : [];
	const outerField = pivot.fields.indexOf(pivot.rowField);
	const innerField = pivot.fields.indexOf(pivot.secondRowField!);
	const columnField = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const layout = pivot.layout ?? 'tabular';
	const compact = layout === 'compact';
	const top =
		layout !== 'tabular' &&
		pivot.subtotals !== false &&
		(pivot.subtotalPosition ?? 'top') === 'top';
	const bottom = pivot.subtotals !== false && !top;
	const rowLabelColumns = compact ? 1 : 2;
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
	const labels = compact ? ['Row Labels'] : [pivot.rowField, pivot.secondRowField!];
	const values = pivotHeaders(pivot, columnItems, labels, label);
	const rows =
		details.size +
		rowItems.length * (Number(layout !== 'tabular') + Number(bottom)) +
		values.length +
		1;
	if (rows * values[0]!.length > 1_000_000) throw new Error('The PivotTable result is too large.');
	const firstDataRow = values.length;
	const indents = values.map(() => 0);
	const rowEntries: NonNullable<PivotGrid['rowEntries']> = [];
	const append = (
		labels: CellValue[],
		numbers: CellValue[],
		entry: NonNullable<PivotGrid['rowEntries']>[number],
		indent = 0,
	) => {
		values.push([...labels, ...numbers]);
		rowEntries.push(entry);
		indents.push(indent);
	};
	const numbers = (keys: string[], records: CellValue[][]): CellValue[] => [
		...columnItems.flatMap((column) =>
			pivotNumbers(pivot, buckets.get(JSON.stringify([...keys, pivotValueKey(column)])) ?? []),
		),
		...pivotNumbers(pivot, records),
	];
	rowItems.forEach((outer, outerIndex) => {
		const key = pivotValueKey(outer);
		if (layout !== 'tabular')
			append(
				[label(outer), ...(compact ? [] : [null])],
				top
					? numbers([key], groups.get(key) ?? [])
					: Array((columnItems.length + 1) * pivotMeasures(pivot).length).fill(null),
				{ indices: [outerIndex] },
			);
		let first = true;
		[...children.get(key)!]
			.sort((a, b) => innerIndices.get(a)! - innerIndices.get(b)!)
			.forEach((innerKey) => {
				const innerIndex = innerIndices.get(innerKey)!;
				const inner = secondRowItems[innerIndex]!;
				const keys = [key, innerKey];
				const records = details.get(JSON.stringify(keys));
				if (!records) return;
				append(
					compact
						? [label(inner)]
						: [
								pivot.repeatLabels || (layout === 'tabular' && first) ? label(outer) : null,
								label(inner),
							],
					numbers(keys, records),
					{ indices: [outerIndex, innerIndex] },
					compact ? 1 : 0,
				);
				first = false;
			});
		if (bottom) {
			const caption = String(label(outer));
			append(
				[`${caption} Total`, ...(compact ? [] : [null])],
				numbers([key], groups.get(key) ?? []),
				{ indices: [outerIndex], type: 'default' },
			);
		}
	});
	append(['Grand Total', ...(compact ? [] : [null])], numbers([], pivot.records), {
		indices: [0],
		type: 'grand',
	});
	return {
		values,
		rowItems,
		secondRowItems,
		columnItems,
		rowEntries,
		firstDataRow,
		rowLabelColumns,
		indents,
	};
}
