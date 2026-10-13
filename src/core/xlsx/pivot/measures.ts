import {
	isCellError,
	type CellValue,
	type PivotAggregate,
	type PivotTable,
	type PivotValueField,
	type PivotValueFieldOptions,
} from '../model';
import { PIVOT_AGGREGATES } from './data';

export const pivotMeasures = (
	pivot: Pick<PivotTable, 'valueField' | 'aggregate' | 'valueCaption' | 'additionalValues'>,
): PivotValueField[] => [
	{ field: pivot.valueField, aggregate: pivot.aggregate, caption: pivot.valueCaption },
	...(pivot.additionalValues ?? []),
];
export function normalizedPivotMeasures(
	values: readonly PivotValueFieldOptions[],
): PivotValueField[] {
	if (!Array.isArray(values) || values.length > 15)
		throw new Error('Choose between one and sixteen PivotTable value fields.');
	return values.map((value) => {
		if (!value || typeof value.field !== 'string')
			throw new Error('Choose valid PivotTable value fields.');
		const aggregate = value.aggregate ?? 'sum';
		if (!PIVOT_AGGREGATES.includes(aggregate))
			throw new Error('Choose a valid PivotTable aggregate.');
		return {
			field: value.field,
			aggregate,
			caption:
				value.caption ?? `${aggregate[0]!.toUpperCase()}${aggregate.slice(1)} of ${value.field}`,
		};
	});
}

/** Every measure uses its own source values, including errors and missing intersections. */
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
export const pivotNumbers = (pivot: PivotTable, records: readonly CellValue[][]): CellValue[] =>
	pivotMeasures(pivot).map((measure) =>
		aggregate(records, pivot.fields.indexOf(measure.field), measure.aggregate),
	);

/** Allocation-free header dimensions, including one totals column for every measure. */
export function pivotHeaderDimensions(
	pivot: Pick<PivotTable, 'columnField' | 'additionalValues'>,
	columnItems: number,
	rowLabels: number,
): { rows: number; width: number } {
	const measures = 1 + (pivot.additionalValues?.length ?? 0);
	return {
		rows: pivot.columnField ? (measures > 1 ? 3 : 2) : 1,
		width:
			rowLabels +
			(pivot.columnField
				? measures === 1
					? Math.max(1, columnItems) + 1
					: (columnItems + 1) * measures
				: measures),
	};
}

/** Values follows the optional column field; native totals have one column per measure. */
export function pivotHeaders(
	pivot: PivotTable,
	columns: readonly CellValue[],
	labels: CellValue[],
	label: (value: CellValue) => CellValue,
): CellValue[][] {
	const measures = pivotMeasures(pivot);
	if (!pivot.columnField) return [[...labels, ...measures.map((measure) => measure.caption)]];
	if (measures.length === 1)
		return [
			[
				pivot.valueCaption,
				...labels.slice(1).map(() => null),
				pivot.layout === 'compact' ? 'Column Labels' : pivot.columnField,
				...columns.slice(1).map(() => null),
				null,
			],
			[...labels, ...columns.map(label), 'Grand Total'],
		];
	const { width } = pivotHeaderDimensions(pivot, columns.length, labels.length);
	const first: CellValue[] = Array(width).fill(null);
	first[labels.length] = pivot.layout === 'compact' ? 'Column Labels' : pivot.columnField;
	if (pivot.layout !== 'compact') first[labels.length + 1] = 'Values';
	return [
		first,
		[
			...labels.map(() => null),
			...columns.flatMap((column) => [label(column), ...measures.slice(1).map(() => null)]),
			...measures.map((measure) => `Total ${measure.caption}`),
		],
		[
			...labels,
			...columns.flatMap(() => measures.map((measure) => measure.caption)),
			...measures.map(() => null),
		],
	];
}
