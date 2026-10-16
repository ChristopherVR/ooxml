import { children, first, NS, type XmlElement } from '../../xml/index';
import {
	isCellError,
	isErrorCode,
	type CellValue,
	type PivotPageFilter,
	type PivotTable,
} from '../model';
import { pivotItems, pivotItemLabel, pivotValueKey } from './data';
import { pivotItemLabeler } from './item-labels';
import type { PivotGrid } from './aggregate';
import { pivotCaptions } from './captions';
import { resolveDateGroups } from './date-groups';
import { applyPivotGroups } from './groups';
import { applyPivotItemFilters } from './item-filters';

export function validatePivotPageFilter(
	pivot: Pick<
		PivotTable,
		| 'pageFilter'
		| 'fields'
		| 'rowField'
		| 'secondRowField'
		| 'columnField'
		| 'valueField'
		| 'additionalValues'
	>,
): void {
	const filter = pivot.pageFilter;
	const extras = Array.isArray(pivot.additionalValues) ? pivot.additionalValues : [];
	if (filter === undefined) return;
	if (
		!filter ||
		typeof filter !== 'object' ||
		Array.isArray(filter) ||
		Object.keys(filter).some((key) => key !== 'field' && key !== 'item') ||
		typeof filter.field !== 'string' ||
		!pivot.fields.includes(filter.field) ||
		[
			pivot.rowField,
			pivot.secondRowField,
			pivot.columnField,
			pivot.valueField,
			...extras.map((v) => v?.field),
		].includes(filter.field)
	)
		throw new Error('Report filters require a separate source field.');
	if (
		filter.item !== undefined &&
		!(
			filter.item === null ||
			typeof filter.item === 'string' ||
			typeof filter.item === 'boolean' ||
			(typeof filter.item === 'number' && Number.isFinite(filter.item)) ||
			(isCellError(filter.item) &&
				isErrorCode(filter.item.error) &&
				Object.keys(filter.item).length === 1)
		)
	)
		throw new Error('Choose a cached literal report filter item.');
}
/** The cache remains complete; only the existing aggregation projection sees selected records. */
export function pivotPageProjection(pivot: PivotTable): PivotTable {
	// The cache fixes the date ranges from every record, so they are resolved before any filter.
	const groups = pivot.dateGroups?.length
		? { ...pivot, dateGroups: resolveDateGroups(pivot) }
		: pivot;
	return applyPivotGroups(applyPivotItemFilters(pagedProjection(groups)));
}
function pagedProjection(pivot: PivotTable): PivotTable {
	const filter = pivot.pageFilter;
	if (!filter) return pivot;
	const labels = new Map<string, string>();
	for (const item of pivotItems(pivot, filter.field)) {
		const label = String(pivotItemLabel(item)).toLowerCase(),
			key = pivotValueKey(item);
		if (label === '(all)' || (labels.has(label) && labels.get(label) !== key))
			throw new Error('Ambiguous report filter item captions are preserved.');
		labels.set(label, key);
	}
	if (filter.item === undefined) return pivot;
	const index = pivot.fields.indexOf(filter.field),
		key = pivotValueKey(filter.item);
	const records = pivot.records.filter((record) => pivotValueKey(record[index] ?? null) === key);
	if (!records.length) throw new Error('The report filter item is missing from the source.');
	if (JSON.stringify(records[0]![index] ?? null) !== JSON.stringify(filter.item))
		throw new Error('Choose an exact cached literal report filter item.');
	return { ...pivot, records };
}
export function prependPivotPage(grid: PivotGrid, pivot: PivotTable): PivotGrid {
	if (!pivot.pageFilter) return grid;
	const width = grid.values[0]!.length;
	return {
		...grid,
		values: [
			[
				pivot.pageFilter.field,
				pivot.pageFilter.item === undefined
					? pivotCaptions(pivot).all
					: pivotItemLabeler(pivot)(pivot.pageFilter.field, pivot.pageFilter.item),
				...Array<CellValue>(width - 2).fill(null),
			],
			Array<CellValue>(width).fill(null),
			...grid.values,
		],
		firstDataRow: grid.firstDataRow + 2,
		...(grid.firstHeaderRow !== undefined ? { firstHeaderRow: grid.firstHeaderRow + 2 } : {}),
		...(grid.indents ? { indents: [0, 0, ...grid.indents] } : {}),
	};
}
const index = (raw: string | null): number =>
	raw !== null && /^\d+$/.test(raw.trim()) && Number.isSafeInteger(Number(raw)) ? Number(raw) : NaN;
export function readPivotPageFilter(
	root: XmlElement,
	fields: string[],
	shared: CellValue[][],
): PivotPageFilter | undefined {
	const list = first(root, 'pageFields', NS.x),
		location = first(root, 'location', NS.x);
	const definitions = children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x);
	if (!list) {
		if (
			definitions.some((n) => n.getAttribute('axis') === 'axisPage') ||
			location?.hasAttribute('rowPageCount') ||
			location?.hasAttribute('colPageCount')
		)
			throw new Error('Unmodeled report filter metadata');
		return;
	}
	if (root.getAttribute('multipleFieldFilters') === '1')
		throw new Error('Multiple report filters are preserved');
	const pages = children(list, 'pageField', NS.x);
	if (
		pages.length !== 1 ||
		list.getAttribute('count') !== '1' ||
		location?.getAttribute('rowPageCount') !== '1' ||
		location.getAttribute('colPageCount') !== '1'
	)
		throw new Error('Only one report filter is modeled');
	const page = pages[0]!,
		field = index(page.getAttribute('fld'));
	if (
		page.getAttribute('hier') !== '-1' ||
		!fields[field] ||
		definitions[field]?.getAttribute('axis') !== 'axisPage' ||
		definitions.filter((n) => n.getAttribute('axis') === 'axisPage').length !== 1
	)
		throw new Error('Unsupported report field');
	const container = first(definitions[field], 'items', NS.x);
	const items = children(container!, 'item', NS.x);
	if (
		!container ||
		index(container.getAttribute('count')) !== items.length ||
		items.filter((n) => n.getAttribute('t') === 'default').length !== 1 ||
		items.at(-1)?.getAttribute('t') !== 'default' ||
		items.at(-1)?.hasAttribute('x')
	)
		throw new Error('Unsupported report field items');
	const data = items.slice(0, -1).map((item) => {
		const i = index(item.getAttribute('x'));
		if (item.hasAttribute('t') || !Number.isInteger(i) || i >= shared[field]!.length)
			throw new Error('Invalid report cache item');
		return i;
	});
	if (new Set(data).size !== data.length || data.length !== shared[field]!.length)
		throw new Error('Incomplete report cache items');
	const item = page.getAttribute('item');
	if (item === null) return { field: fields[field]! };
	const selected = index(item);
	if (!Number.isInteger(selected) || selected >= data.length)
		throw new Error('Invalid report item selection');
	return { field: fields[field]!, item: shared[field]![data[selected]!]! };
}
/** Stable public item choices retain numeric, boolean, blank and error cache types. */
export function pivotPageItems(
	pivot: Pick<PivotTable, 'fields' | 'records'>,
	field: string,
): CellValue[] {
	return pivotItems(pivot, field);
}
