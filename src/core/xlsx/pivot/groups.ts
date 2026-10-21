import type { CellValue, PivotFieldGroup, PivotTable } from '../model';
import { isCellError } from '../model';
import { pivotColumnFields, pivotRowFields } from './axes';
import { derivedDateLabels, expandDateGroups, validateDateGroups } from './date-groups';

const MAX_BUCKETS = 10_000;

/** The bucket count of a number group: `by` wide from `start`, the last one closed at `end`. */
export function groupBucketCount(group: Pick<PivotFieldGroup, 'start' | 'end' | 'by'>): number {
	return Math.max(1, Math.ceil((group.end - group.start) / group.by));
}

/**
 * Group item labels in Excel's order: `<start`, one `lo-hi` label per bucket (the last runs to
 * `end`), `>end`. The two outer items always exist, even when no record falls in them.
 */
export function numberGroupLabels(group: Pick<PivotFieldGroup, 'start' | 'end' | 'by'>): string[] {
	const count = groupBucketCount(group);
	const labels = [`<${group.start}`];
	for (let k = 0; k < count; k++) {
		const lo = group.start + k * group.by;
		const hi = k === count - 1 ? group.end : lo + group.by - 1;
		labels.push(`${lo}-${hi}`);
	}
	labels.push(`>${group.end}`);
	return labels;
}

/** The label of the bucket a number falls in. */
export function numberGroupLabel(group: PivotFieldGroup, value: number): string {
	if (value < group.start) return `<${group.start}`;
	if (value > group.end) return `>${group.end}`;
	const count = groupBucketCount(group);
	const k = Math.min(count - 1, Math.floor((value - group.start) / group.by));
	return numberGroupLabels(group)[k + 1]!;
}

/** Whether the values of `field` can be grouped (every record holds a number). */
export function validateFieldGroups(
	pivot: Pick<
		PivotTable,
		| 'fields'
		| 'records'
		| 'fieldGroups'
		| 'pageFilter'
		| 'itemFilters'
		| 'dateFields'
		| 'dateFilter'
		| 'dateGroups'
		| 'date1904'
		| 'valueField'
		| 'additionalValues'
	>,
): void {
	validateDateGroups(pivot);
	const groups = pivot.fieldGroups;
	if (groups === undefined) return;
	if (!Array.isArray(groups)) throw new Error('Choose a valid number group.');
	const seen = new Set<string>();
	for (const group of groups) {
		const index = pivot.fields.indexOf(group?.field as string);
		if (
			!group ||
			typeof group !== 'object' ||
			index < 0 ||
			seen.has(group.field) ||
			![group.start, group.end, group.by].every(Number.isInteger) ||
			group.by < 1 ||
			group.start >= group.end ||
			groupBucketCount(group) > MAX_BUCKETS
		)
			throw new Error('Choose a valid number group.');
		seen.add(group.field);
		if (
			pivot.pageFilter?.field === group.field ||
			pivot.valueField === group.field ||
			(pivot.additionalValues ?? []).some((measure) => measure.field === group.field) ||
			(pivot.dateFields ?? []).includes(group.field) ||
			(pivot.itemFilters ?? []).some((filter) => filter.field === group.field)
		)
			throw new Error('A grouped field cannot also be a filter, date field or value field.');
		for (const record of pivot.records) {
			const value = record[index];
			if (typeof value !== 'number' || !Number.isFinite(value))
				throw new Error('Only fields that contain numbers can be grouped.');
		}
	}
}

/** The records with every grouped field replaced by its bucket label (the cache is unchanged). */
export function applyPivotGroups(source: PivotTable): PivotTable {
	const pivot = expandDateGroups(source);
	const groups = (pivot.fieldGroups ?? []).filter((group) => pivot.fields.includes(group.field));
	if (!groups.length) return pivot;
	const targets = groups.map((group) => ({ group, index: pivot.fields.indexOf(group.field) }));
	return {
		...pivot,
		records: pivot.records.map((record) => {
			const next: CellValue[] = [...record];
			for (const { group, index } of targets) {
				const value = record[index];
				if (typeof value === 'number') next[index] = numberGroupLabel(group, value);
			}
			return next;
		}),
	};
}

/** The group of a field, if any. */
export const groupOf = (
	pivot: { fieldGroups?: PivotFieldGroup[] | undefined },
	field: string,
): PivotFieldGroup | undefined => pivot.fieldGroups?.find((group) => group.field === field);

/** Sort rank of a group label (its index in the group's item list); labels from elsewhere last. */
export function groupRank(labels: readonly string[], value: CellValue): number {
	if (typeof value !== 'string' || isCellError(value)) return Number.MAX_SAFE_INTEGER;
	const at = labels.indexOf(value);
	return at < 0 ? Number.MAX_SAFE_INTEGER : at;
}

/** The ordered group items of a number-grouped or derived date field, if it is one. */
export function groupLabelsOf(
	pivot: Pick<PivotTable, 'fields' | 'records'> & {
		fieldGroups?: PivotTable['fieldGroups'];
		dateGroups?: PivotTable['dateGroups'];
		date1904?: boolean | undefined;
	},
	field: string,
): string[] | undefined {
	const group = groupOf(pivot, field);
	return group ? numberGroupLabels(group) : derivedDateLabels(pivot, field);
}

/** Axis fields whose every value is an integer, so they can be grouped by number ranges. */
export function pivotGroupableFields(
	pivot: Pick<
		PivotTable,
		| 'fields'
		| 'records'
		| 'rowField'
		| 'secondRowField'
		| 'moreRowFields'
		| 'columnField'
		| 'moreColumnFields'
		| 'dateFields'
	>,
): string[] {
	return [...pivotRowFields(pivot), ...pivotColumnFields(pivot)].filter(
		(field): field is string => {
			if (!field || (pivot.dateFields ?? []).includes(field)) return false;
			const index = pivot.fields.indexOf(field);
			return (
				pivot.records.length > 0 && pivot.records.every((record) => Number.isInteger(record[index]))
			);
		},
	);
}

/** The group the dialog offers: the data range split into about ten buckets. */
export function defaultFieldGroup(
	pivot: Pick<PivotTable, 'fields' | 'records' | 'fieldGroups'>,
	field: string,
): PivotFieldGroup {
	const existing = groupOf(pivot, field);
	if (existing) return { ...existing };
	const index = pivot.fields.indexOf(field);
	// The extremes in a loop: a spread passes one argument per record and overflows on large pivots.
	let start = Infinity;
	let end = -Infinity;
	for (const record of pivot.records) {
		const value = Number(record[index]);
		start = Math.min(start, value);
		end = Math.max(end, value);
	}
	return {
		field,
		start,
		end: Math.max(end, start + 1),
		by: Math.max(1, Math.ceil((end - start + 1) / 10)),
	};
}
