import type { PivotDateGroup, PivotTable } from '../model';
import { pivotColumnFields, pivotRowFields } from './axes';
import { dateGroupFieldName, dateGroupUnits } from './date-groups';

type AxisPivot = Pick<
	PivotTable,
	| 'rowField'
	| 'secondRowField'
	| 'moreRowFields'
	| 'columnField'
	| 'moreColumnFields'
	| 'dateGroups'
	| 'fieldSubtotals'
>;

/** The axis names a group stands for: coarsest first (`Years`, `Quarters`, `Months`). */
export const dateGroupAxisNames = (
	group: Pick<PivotDateGroup, 'field' | 'by' | 'localized'>,
): string[] =>
	dateGroupUnits(group)
		.reverse()
		.map((unit) => dateGroupFieldName(unit, group.field, group));

function replace(
	list: string[],
	names: readonly string[],
	replacement: readonly string[],
): string[] {
	const at = list.findIndex((name) => names.includes(name));
	if (at < 0) return list;
	const rest = list.filter((name) => !names.includes(name));
	const before = list.slice(0, at).filter((name) => !names.includes(name)).length;
	return [
		...rest.slice(0, before),
		...replacement.filter((name) => !rest.includes(name)),
		...rest.slice(before),
	];
}

/**
 * Keeps the axes consistent with the date groups, like Excel's Group and Ungroup do: a grouped
 * date field is replaced on the axes by its group fields (coarsest outermost), a changed group
 * swaps in the new fields, and ungrouping puts the date field back. Mutates `next`.
 */
export function reconcileDateAxes(
	next: AxisPivot,
	previous?: Pick<PivotTable, 'dateGroups'>,
): void {
	let rows = pivotRowFields(next);
	let columns = pivotColumnFields(next);
	const swap = (names: readonly string[], replacement: readonly string[]) => {
		rows = replace(rows, names, replacement);
		columns = replace(columns, names, replacement);
	};
	const groups = next.dateGroups ?? [];
	for (const old of previous?.dateGroups ?? []) {
		const current = groups.find((group) => group.field === old.field);
		// The same units keep whatever subset of their fields the axes use.
		if (current && dateGroupAxisNames(current).join() === dateGroupAxisNames(old).join()) continue;
		swap(dateGroupAxisNames(old), current ? dateGroupAxisNames(current) : [old.field]);
	}
	for (const group of groups) swap([group.field], dateGroupAxisNames(group));
	next.rowField = rows[0]!;
	if (rows[1]) next.secondRowField = rows[1];
	else delete next.secondRowField;
	if (rows.length > 2) next.moreRowFields = rows.slice(2);
	else delete next.moreRowFields;
	if (columns[0]) next.columnField = columns[0];
	else delete next.columnField;
	if (columns.length > 1) next.moreColumnFields = columns.slice(1);
	else delete next.moreColumnFields;
	if (next.fieldSubtotals) {
		const kept = next.fieldSubtotals.filter(
			(entry) => rows.includes(entry.field) || columns.includes(entry.field),
		);
		if (kept.length) next.fieldSubtotals = kept;
		else delete next.fieldSubtotals;
	}
}
