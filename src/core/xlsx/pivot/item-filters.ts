import type { CellValue, PivotItemFilter, PivotTable } from '../model';
import { pivotItems, pivotValueKey } from './data';

type FilterFields = Pick<PivotTable, 'fields' | 'records'> & {
	itemFilters?: PivotTable['itemFilters'];
};

/** The cached-item keys hidden from `field`, empty when it has no item filter. */
export function hiddenItemKeys(pivot: Pick<PivotTable, 'itemFilters'>, field: string): Set<string> {
	const filter = pivot.itemFilters?.find((entry) => entry.field === field);
	return new Set((filter?.hidden ?? []).map(pivotValueKey));
}

/** Fields that carry at least one hidden item. */
export function filteredFields(pivot: Pick<PivotTable, 'itemFilters'>): Set<string> {
	return new Set(
		(pivot.itemFilters ?? []).filter((entry) => entry.hidden.length).map((entry) => entry.field),
	);
}

/** The pivot restricted to the records whose filtered fields show a visible item. */
export function applyPivotItemFilters(pivot: PivotTable): PivotTable {
	const active = (pivot.itemFilters ?? []).filter((entry) => entry.hidden.length);
	if (!active.length) return pivot;
	const tests = active.map((entry) => ({
		index: pivot.fields.indexOf(entry.field),
		hidden: new Set(entry.hidden.map(pivotValueKey)),
	}));
	return {
		...pivot,
		records: pivot.records.filter((record) =>
			tests.every((test) => !test.hidden.has(pivotValueKey(record[test.index] ?? null))),
		),
	};
}

/** Items of one field that still have at least one record under the other filters. */
export function itemsWithData(pivot: PivotTable, field: string): Set<string> {
	const others: PivotItemFilter[] = (pivot.itemFilters ?? []).filter((e) => e.field !== field);
	const scoped = applyPivotItemFilters({ ...pivot, itemFilters: others });
	const index = pivot.fields.indexOf(field);
	return new Set(scoped.records.map((record) => pivotValueKey(record[index] ?? null)));
}

export function validatePivotItemFilters(pivot: FilterFields & { pageFilter?: { field: string } }) {
	const filters = pivot.itemFilters;
	if (filters === undefined) return;
	if (!Array.isArray(filters)) throw new Error('Choose valid PivotTable item filters.');
	const seen = new Set<string>();
	for (const filter of filters) {
		if (
			!filter ||
			typeof filter !== 'object' ||
			typeof filter.field !== 'string' ||
			!Array.isArray(filter.hidden) ||
			!pivot.fields.includes(filter.field) ||
			seen.has(filter.field) ||
			pivot.pageFilter?.field === filter.field
		)
			throw new Error('Choose valid PivotTable item filters.');
		seen.add(filter.field);
		const items = new Map(pivotItems(pivot, filter.field).map((v) => [pivotValueKey(v), v]));
		const keys = filter.hidden.map(pivotValueKey);
		if (new Set(keys).size !== keys.length || keys.some((key) => !items.has(key)))
			throw new Error('A hidden PivotTable item is not in the source.');
		if (keys.length && keys.length >= items.size)
			throw new Error('At least one PivotTable item must stay visible.');
	}
}

/** Drops hidden items the refreshed source no longer has (and filters left empty). */
export function pruneItemFilters(pivot: FilterFields): PivotItemFilter[] | undefined {
	if (!pivot.itemFilters) return undefined;
	const next: PivotItemFilter[] = [];
	for (const filter of pivot.itemFilters) {
		if (!pivot.fields.includes(filter.field)) continue;
		const present = new Map(pivotItems(pivot, filter.field).map((v) => [pivotValueKey(v), v]));
		const hidden = filter.hidden.filter((value: CellValue) => present.has(pivotValueKey(value)));
		if (hidden.length && hidden.length < present.size) next.push({ field: filter.field, hidden });
	}
	return next.length ? next : undefined;
}
