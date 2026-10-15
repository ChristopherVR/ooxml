import type { PivotTable } from '../model';

/** The row fields, outermost first: the first, second and any further ones. */
export const pivotRowFields = (
	pivot: Pick<PivotTable, 'rowField' | 'secondRowField' | 'moreRowFields'>,
): string[] => [
	pivot.rowField,
	...(pivot.secondRowField ? [pivot.secondRowField] : []),
	...(pivot.secondRowField ? (pivot.moreRowFields ?? []) : []),
];

/**
 * Where Values sit among the row fields of a Values-on-rows report: the number of row fields
 * that come before it (0 is first, the field count is last). -1 means Values on columns.
 */
export function pivotValuesIndex(
	pivot: Pick<
		PivotTable,
		'valuesAxis' | 'valuesPosition' | 'rowField' | 'secondRowField' | 'moreRowFields'
	>,
): number {
	if (pivot.valuesAxis !== 'rows') return -1;
	const count = pivotRowFields(pivot).length;
	const at = pivot.valuesPosition;
	if (typeof at === 'number') return Math.min(Math.max(Math.trunc(at), 0), count);
	return at === 'first' ? 0 : at === 'between' ? Math.min(1, count) : count;
}

/** The column fields, outermost first. */
export const pivotColumnFields = (
	pivot: Pick<PivotTable, 'columnField' | 'moreColumnFields'>,
): string[] => [
	...(pivot.columnField ? [pivot.columnField] : []),
	...(pivot.columnField ? (pivot.moreColumnFields ?? []) : []),
];
