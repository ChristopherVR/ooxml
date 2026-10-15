import type { PivotAggregate, PivotFieldSubtotals, PivotTable } from '../model';
import { PIVOT_AGGREGATES } from './data';

/**
 * How Excel names each subtotal function: the `pivotField` flag, the `item/@t` of the subtotal
 * row and the label suffix (`Fruit Sum`). Count and Count Numbers share the label "Count".
 */
export const PIVOT_SUBTOTALS: Record<
	PivotAggregate,
	{ attr: string; item: string; label: string }
> = {
	sum: { attr: 'sumSubtotal', item: 'sum', label: 'Sum' },
	count: { attr: 'countASubtotal', item: 'countA', label: 'Count' },
	average: { attr: 'avgSubtotal', item: 'avg', label: 'Average' },
	max: { attr: 'maxSubtotal', item: 'max', label: 'Max' },
	min: { attr: 'minSubtotal', item: 'min', label: 'Min' },
	product: { attr: 'productSubtotal', item: 'product', label: 'Product' },
	countNums: { attr: 'countSubtotal', item: 'count', label: 'Count' },
	stdDev: { attr: 'stdDevSubtotal', item: 'stdDev', label: 'StdDev' },
	stdDevp: { attr: 'stdDevPSubtotal', item: 'stdDevP', label: 'StdDevp' },
	var: { attr: 'varSubtotal', item: 'var', label: 'Var' },
	varp: { attr: 'varPSubtotal', item: 'varP', label: 'Varp' },
};

/** The order Excel lists subtotal rows in, whatever order they were chosen in. */
export const PIVOT_SUBTOTAL_ORDER: readonly PivotAggregate[] = [
	'sum',
	'count',
	'average',
	'max',
	'min',
	'product',
	'countNums',
	'stdDev',
	'stdDevp',
	'var',
	'varp',
];

/** The function of a subtotal row's `item/@t` (anything but `default`), for reading. */
export const subtotalFromItem = (type: string | null): PivotAggregate | undefined =>
	PIVOT_SUBTOTAL_ORDER.find((fn) => PIVOT_SUBTOTALS[fn].item === type);

/**
 * The subtotals a row field shows: `'auto'` (a Total row, or the group row's value when
 * subtotals sit above), or the subtotal functions in Excel's order (empty means none).
 */
export function fieldSubtotalFunctions(
	pivot: Pick<PivotTable, 'rowField' | 'subtotals' | 'fieldSubtotals'>,
	field: string,
): 'auto' | PivotAggregate[] {
	const entry = pivot.fieldSubtotals?.find((item) => item.field === field);
	if (entry) return PIVOT_SUBTOTAL_ORDER.filter((fn) => entry.functions.includes(fn));
	return field === pivot.rowField && pivot.subtotals === false ? [] : 'auto';
}

/** The entries to store: functions are unique and in Excel's order. */
export function normalizedFieldSubtotals(
	entries: readonly PivotFieldSubtotals[],
): PivotFieldSubtotals[] {
	return entries.map((entry) => {
		if (
			!entry ||
			!Array.isArray(entry.functions) ||
			entry.functions.some((fn) => !PIVOT_SUBTOTAL_ORDER.includes(fn))
		)
			throw new Error('Choose valid field subtotals.');
		return {
			field: entry.field,
			functions: PIVOT_SUBTOTAL_ORDER.filter((fn) => entry.functions.includes(fn)),
		};
	});
}

export function validateFieldSubtotals(
	pivot: Pick<PivotTable, 'fieldSubtotals' | 'valuesAxis'>,
	rowFields: readonly string[],
): void {
	const entries = pivot.fieldSubtotals;
	if (entries === undefined) return;
	if (!Array.isArray(entries)) throw new Error('Choose valid field subtotals.');
	if (entries.length && pivot.valuesAxis === 'rows')
		throw new Error('Field subtotals require Values on columns.');
	const seen = new Set<string>();
	for (const entry of entries) {
		if (
			!entry ||
			typeof entry !== 'object' ||
			!rowFields.includes(entry.field) ||
			seen.has(entry.field) ||
			!Array.isArray(entry.functions) ||
			entry.functions.some((fn) => !PIVOT_AGGREGATES.includes(fn)) ||
			new Set(entry.functions).size !== entry.functions.length
		)
			throw new Error('Choose valid field subtotals.');
		seen.add(entry.field);
	}
}
