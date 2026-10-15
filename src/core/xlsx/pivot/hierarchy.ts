import type { CellValue, PivotTable } from '../model';
import type { PivotGrid } from './aggregate';
import { pivotItemLabel, pivotItems, pivotValueKey } from './data';
import { appendFirstPivotValues } from './hierarchy-first';

import {
	pivotHeaderDimensions,
	pivotHeaders,
	pivotNumbers,
	pivotMeasures,
	pivotMeasureNumber,
} from './measures';
const label = pivotItemLabel;

/**
 * Values on rows with two row fields (the measures sit last, between or before them). Values on
 * columns use `nestedPivotGrid`, which handles any number of row fields.
 */
export function hierarchicalPivotGrid(pivot: PivotTable): PivotGrid {
	const rowItems = pivotItems(pivot, pivot.rowField);
	const secondRowItems = pivotItems(pivot, pivot.secondRowField!);
	const columnItems = pivot.columnField ? pivotItems(pivot, pivot.columnField) : [];
	const outerField = pivot.fields.indexOf(pivot.rowField);
	const innerField = pivot.fields.indexOf(pivot.secondRowField!);
	const columnField = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const layout = pivot.layout ?? 'tabular';
	const compact = layout === 'compact';
	const firstValues = pivot.valuesPosition === 'first';
	const between = pivot.valuesPosition === 'between';
	const measures = pivotMeasures(pivot);
	// Last/between row Values put subtotals below groups; first Values honor top outside tabular.
	const top =
		firstValues &&
		layout !== 'tabular' &&
		pivot.subtotals !== false &&
		(pivot.subtotalPosition ?? 'top') === 'top';
	const bottom = pivot.subtotals !== false && !top;
	const rowLabelColumns = compact ? 1 : 3;
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
	const dimensions = pivotHeaderDimensions(pivot, columnItems.length, rowLabelColumns);
	const rows =
		(details.size + rowItems.length * Number(bottom) + 1) * measures.length +
		rowItems.length * Number(layout !== 'tabular') * (firstValues ? measures.length : 1) +
		(firstValues ? measures.length : between ? rowItems.length * measures.length : details.size) *
			Number(layout !== 'tabular') +
		dimensions.rows;
	if (rows * dimensions.width > 1_000_000) throw new Error('The PivotTable result is too large.');
	const labels = compact ? ['Row Labels'] : [pivot.rowField, pivot.secondRowField!];
	const values = pivotHeaders(pivot, columnItems, labels, label);
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
	const bands = (keys: string[], records: CellValue[][]): CellValue[][] => [
		...columnItems.map((column) =>
			pivotNumbers(pivot, buckets.get(JSON.stringify([...keys, pivotValueKey(column)])) ?? []),
		),
		pivotNumbers(pivot, records),
	];
	const numbers = (keys: string[], records: CellValue[][]) => bands(keys, records).flat();
	const appendMeasures = (
		labels: CellValue[],
		keys: string[],
		records: CellValue[][],
		indices: number[],
		type?: 'default' | 'grand',
	) => {
		const values = bands(keys, records);
		measures.forEach((measure, m) =>
			append(
				type
					? [
							`${type === 'grand' ? 'Total' : String(labels[0])} ${measure.caption}`,
							...Array<CellValue>(rowLabelColumns - 1).fill(null),
						]
					: compact
						? [measure.caption]
						: [
								...labels.map((label) =>
									(layout === 'tabular' && m === 0) || pivot.repeatLabels ? label : null,
								),
								measure.caption,
							],
				values.map((band) => band[m] ?? null),
				{ indices: type ? indices : [...indices, m], measure: m, ...(type ? { type } : {}) },
				compact && !type ? 2 : 0,
			),
		);
	};
	if (firstValues)
		appendFirstPivotValues(pivot, {
			rowItems,
			secondRowItems,
			columnItems,
			children,
			innerIndices,
			groups,
			details,
			buckets,
			width: dimensions.width,
			rowLabels: rowLabelColumns,
			top,
			bottom,
			append,
		});
	else
		rowItems.forEach((outer, outerIndex) => {
			const key = pivotValueKey(outer);
			if (layout !== 'tabular')
				append(
					[label(outer), ...Array<CellValue>(rowLabelColumns - 1).fill(null)],
					top
						? numbers([key], groups.get(key) ?? [])
						: Array(dimensions.width - rowLabelColumns).fill(null),
					{ indices: [outerIndex] },
				);
			let first = true;
			const innerKeys = [...children.get(key)!].sort(
				(a, b) => innerIndices.get(a)! - innerIndices.get(b)!,
			);
			if (between)
				measures.forEach((measure, m) => {
					if (layout !== 'tabular')
						append(
							compact
								? [measure.caption]
								: [pivot.repeatLabels ? label(outer) : null, measure.caption, null],
							Array<CellValue>(dimensions.width - rowLabelColumns).fill(null),
							{ indices: [outerIndex, m], measure: m },
							compact ? 1 : 0,
						);
					innerKeys.forEach((innerKey, child) => {
						const innerIndex = innerIndices.get(innerKey)!;
						const keys = [key, innerKey];
						const records = details.get(JSON.stringify(keys))!;
						append(
							compact
								? [label(secondRowItems[innerIndex]!)]
								: [
										pivot.repeatLabels || (layout === 'tabular' && m === 0 && child === 0)
											? label(outer)
											: null,
										layout === 'tabular' && child === 0 ? measure.caption : null,
										label(secondRowItems[innerIndex]!),
									],
							[
								...columnItems.map((column) =>
									pivotMeasureNumber(
										pivot,
										buckets.get(JSON.stringify([...keys, pivotValueKey(column)])) ?? [],
										measure,
									),
								),
								pivotMeasureNumber(pivot, records, measure),
							],
							{ indices: [outerIndex, m, innerIndex], measure: m },
							compact ? 2 : 0,
						);
					});
				});
			else
				innerKeys.forEach((innerKey) => {
					const innerIndex = innerIndices.get(innerKey)!;
					const inner = secondRowItems[innerIndex]!;
					const keys = [key, innerKey];
					const records = details.get(JSON.stringify(keys));
					if (!records) return;
					if (layout !== 'tabular')
						append(
							compact
								? [label(inner)]
								: [pivot.repeatLabels ? label(outer) : null, label(inner), null],
							Array<CellValue>(dimensions.width - rowLabelColumns).fill(null),
							{ indices: [outerIndex, innerIndex] },
							compact ? 1 : 0,
						);
					appendMeasures(
						[pivot.repeatLabels || first ? label(outer) : null, label(inner)],
						keys,
						records,
						[outerIndex, innerIndex],
					);
					first = false;
				});
			if (bottom)
				appendMeasures([label(outer)], [key], groups.get(key) ?? [], [outerIndex], 'default');
		});
	appendMeasures([], [], pivot.records, [0], 'grand');
	return {
		values,
		rowItems,
		secondRowItems,
		columnItems,
		rowEntries,
		firstDataRow,
		rowLabelColumns,
		indents,
		firstHeaderRow: 1,
	};
}
