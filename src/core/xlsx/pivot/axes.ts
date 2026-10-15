import type { PivotTable } from '../model';

/** The row fields, outermost first: the first, second and any further ones. */
export const pivotRowFields = (
	pivot: Pick<PivotTable, 'rowField' | 'secondRowField' | 'moreRowFields'>,
): string[] => [
	pivot.rowField,
	...(pivot.secondRowField ? [pivot.secondRowField] : []),
	...(pivot.secondRowField ? (pivot.moreRowFields ?? []) : []),
];

/** The column fields, outermost first. */
export const pivotColumnFields = (
	pivot: Pick<PivotTable, 'columnField' | 'moreColumnFields'>,
): string[] => [
	...(pivot.columnField ? [pivot.columnField] : []),
	...(pivot.columnField ? (pivot.moreColumnFields ?? []) : []),
];
