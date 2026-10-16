import {
	isCellError,
	type CellValue,
	type PivotAggregate,
	type PivotTable,
	type PivotValueField,
	type PivotValueFieldOptions,
} from '../model';
import { pivotMeasureName, type CaptionHolder } from './captions';
import { PIVOT_AGGREGATES } from './data';

export const pivotMeasures = (
	pivot: Pick<PivotTable, 'valueField' | 'aggregate' | 'valueCaption' | 'additionalValues'>,
): PivotValueField[] => [
	{ field: pivot.valueField, aggregate: pivot.aggregate, caption: pivot.valueCaption },
	...(pivot.additionalValues ?? []),
];
export function normalizedPivotMeasures(
	values: readonly PivotValueFieldOptions[],
	pivot?: CaptionHolder,
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
			caption: value.caption ?? pivotMeasureName(aggregate, value.field, pivot),
		};
	});
}

const DIV0: CellValue = { error: '#DIV/0!' };

/** Sample (n - 1) or population (n) variance of at least one number, two-pass for accuracy. */
function variance(numbers: readonly number[], sample: boolean): CellValue {
	const n = numbers.length;
	if (n < (sample ? 2 : 1)) return DIV0;
	const mean = numbers.reduce((total, value) => total + value, 0) / n;
	const squares = numbers.reduce((total, value) => total + (value - mean) ** 2, 0);
	return squares / (sample ? n - 1 : n);
}

/** Every measure uses its own source values, including errors and missing intersections. */
export function aggregateValues(
	records: readonly CellValue[][],
	field: number,
	type: PivotAggregate,
): CellValue {
	if (!records.length) return null;
	const values = records.map((record) => record[field] ?? null);
	if (type === 'count') return values.filter((value) => value !== null && value !== '').length;
	if (type === 'countNums') return values.filter((value) => typeof value === 'number').length;
	const error = values.find(isCellError);
	if (error) return error;
	const numbers = values.filter((value): value is number => typeof value === 'number');
	if (type === 'var' || type === 'varp' || type === 'stdDev' || type === 'stdDevp') {
		const result = variance(numbers, type === 'var' || type === 'stdDev');
		return typeof result === 'number' && (type === 'stdDev' || type === 'stdDevp')
			? Math.sqrt(result)
			: result;
	}
	if (!numbers.length) return type === 'average' ? DIV0 : 0;
	if (type === 'min') return numbers.reduce((a, b) => Math.min(a, b));
	if (type === 'max') return numbers.reduce((a, b) => Math.max(a, b));
	if (type === 'product') return numbers.reduce((total, value) => total * value, 1);
	const sum = numbers.reduce((total, value) => total + value, 0);
	return type === 'average' ? sum / numbers.length : sum;
}
export const pivotMeasureNumber = (
	pivot: Pick<PivotTable, 'fields'>,
	records: readonly CellValue[][],
	measure: PivotValueField,
): CellValue => aggregateValues(records, pivot.fields.indexOf(measure.field), measure.aggregate);
export const pivotNumbers = (pivot: PivotTable, records: readonly CellValue[][]): CellValue[] =>
	pivotMeasures(pivot).map((measure) => pivotMeasureNumber(pivot, records, measure));
