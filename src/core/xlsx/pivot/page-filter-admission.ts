import { children, elements, first, NS, type XmlElement } from '../../xml/index';
import type { CellValue, PivotTable } from '../model';
import { pivotItems, pivotRowFields, pivotValueKey } from './data';
import { pivotPageProjection } from './page-filter';
import { pivotGrid, type PivotRowEntry } from './aggregate';
import { PIVOT_SUBTOTALS } from './subtotals';
import { remapPivotRowEntries } from './values-axis';
import { booleanAttribute } from '../../drawingml/dom';
const integer = (value: string | null, fallback?: number) =>
	value === null
		? (fallback ?? NaN)
		: /^\d+$/.test(value.trim()) && Number.isSafeInteger(Number(value))
			? Number(value)
			: NaN;
/** Page profiles retain every axis field item; visible row/column indexes refer to that list. */
export function assertPageAxisItems(
	root: XmlElement,
	pivot: PivotTable,
	shared: CellValue[][],
	cache: XmlElement,
	recordRoot: XmlElement,
): void {
	if (!pivot.pageFilter) return;
	const literals = (node: XmlElement) => {
		const raw = node.getAttribute('v');
		if (
			(node.localName === 'x' && (raw === null || !/^\d+$/.test(raw.trim()))) ||
			(node.localName === 'b' && !['0', '1'].includes(raw ?? '')) ||
			(node.localName === 'n' &&
				(raw === null || !/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw.trim())))
		)
			throw new Error('Unmodeled report cache literal');
		for (const child of elements(node)) literals(child);
	};
	literals(cache);
	literals(recordRoot);
	const cacheFields = children(first(cache, 'cacheFields', NS.x)!, 'cacheField', NS.x);
	if (pivot.columnField) {
		const column = children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x)[
			pivot.fields.indexOf(pivot.columnField)
		];
		const list = first(column, 'extLst', NS.x);
		const extensions = list ? children(list, 'ext', NS.x) : [];
		if (
			extensions.some((extension) =>
				children(
					extension,
					'pivotField',
					'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
				).some((field) => {
					const raw = field.getAttribute('fillDownLabels');
					if (raw !== null && !['0', '1', 'false', 'true'].includes(raw.trim()))
						throw new Error('Invalid report column label repetition');
					return booleanAttribute(field, 'fillDownLabels') === true;
				}),
			)
		)
			throw new Error('Repeated report column labels are preserved');
	}
	if (
		integer(first(cache, 'cacheFields', NS.x)?.getAttribute('count') ?? null) !==
			pivot.fields.length ||
		integer(cache.getAttribute('recordCount')) !== pivot.records.length ||
		integer(recordRoot.getAttribute('count')) !== pivot.records.length
	)
		throw new Error('Inconsistent report cache counts');
	for (const name of [...pivotRowFields(pivot), pivot.columnField, pivot.pageFilter.field].filter(
		Boolean,
	)) {
		const field = pivot.fields.indexOf(name!);
		if (
			integer(first(cacheFields[field], 'sharedItems', NS.x)?.getAttribute('count') ?? null) !==
			shared[field]!.length
		)
			throw new Error('Inconsistent report shared-item counts');
	}
	for (const name of ['rowFields', 'colFields']) {
		const list = first(root, name, NS.x);
		if (list && integer(list.getAttribute('count')) !== children(list, 'field', NS.x).length)
			throw new Error('Inconsistent report axis counts');
	}
	const projected = pivotPageProjection(pivot),
		fields = children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x);
	const axisValues = (name: string, axis: string) => {
		const f = pivot.fields.indexOf(name),
			node = fields[f]!,
			list = first(node, 'items', NS.x),
			items = children(list!, 'item', NS.x),
			count = shared[f]!.length;
		if (node.getAttribute('axis') !== axis) throw new Error('Inconsistent report axis field');
		const defaults = node.getAttribute('defaultSubtotal') !== '0';
		if (
			!list ||
			integer(list.getAttribute('count')) !== items.length ||
			items.length !== count + Number(defaults)
		)
			throw new Error('Incomplete page-axis field items');
		const data = items.slice(0, count).map((item) => {
			const i = integer(item.getAttribute('x'));
			if (item.hasAttribute('t') || !Number.isInteger(i) || i >= count)
				throw new Error('Invalid page-axis cache item');
			return shared[f]![i]!;
		});
		if (
			new Set(data.map(pivotValueKey)).size !== count ||
			(defaults &&
				(items.at(-1)?.getAttribute('t') !== 'default' || items.at(-1)?.hasAttribute('x')))
		)
			throw new Error('Invalid page-axis default item');
		const indices = new Map(data.map((v, i) => [pivotValueKey(v), i]));
		return pivotItems(projected, name).map((v) => indices.get(pivotValueKey(v))!);
	};
	const read = (name: string) => {
		const list = first(root, name, NS.x),
			items = children(list!, 'i', NS.x);
		if (!list || integer(list.getAttribute('count')) !== items.length)
			throw new Error('Invalid page-axis item count');
		return items;
	};
	const rowAxes = pivotRowFields(pivot).map((name) => axisValues(name, 'axisRow')),
		rowItems = read('rowItems');
	const grid = pivotGrid(pivot);
	const entries: PivotRowEntry[] = grid.rowEntries
		? remapPivotRowEntries(grid.rowEntries, rowAxes)
		: [
				...rowAxes[0]!.map((index) => ({ indices: [index] })),
				{ indices: [0], type: 'grand' as const },
			];
	if (rowItems.length !== entries.length) throw new Error('Invalid page row items');
	let previousRow: number[] = [];
	rowItems.forEach((item, i) => {
		const expected = entries[i]!;
		const xs = children(item, 'x', NS.x).map((x) => integer(x.getAttribute('v'), 0)),
			prefix = integer(item.getAttribute('r'), 0);
		const type =
			expected.type === undefined || expected.type === 'default' || expected.type === 'grand'
				? expected.type
				: PIVOT_SUBTOTALS[expected.type].item;
		if (
			item.getAttribute('t') !== (type ?? null) ||
			!Number.isInteger(prefix) ||
			prefix < 0 ||
			prefix > previousRow.length ||
			prefix >= expected.indices.length ||
			(expected.type === 'grand' && prefix !== 0) ||
			integer(item.getAttribute('i'), 0) !== 0 ||
			JSON.stringify([...previousRow.slice(0, prefix), ...xs]) !== JSON.stringify(expected.indices)
		)
			throw new Error('Unmodeled page row item');
		previousRow = expected.type === 'grand' ? [] : expected.indices;
	});
	const measures = 1 + (pivot.additionalValues?.length ?? 0),
		columns = pivot.columnField ? axisValues(pivot.columnField, 'axisCol') : undefined,
		expected: number[][] = [];
	if (columns)
		for (const col of columns)
			for (let m = 0; m < measures; m++) expected.push(measures > 1 ? [col, m] : [col]);
	else for (let m = 0; m < measures; m++) expected.push([m]);
	const colItems = read('colItems');
	if (colItems.length !== expected.length + (columns ? measures : 0))
		throw new Error('Invalid page column items');
	let previous: number[] = [];
	colItems.forEach((item, i) => {
		const grand = i >= expected.length,
			measure = grand ? i - expected.length : i % measures;
		const xs = children(item, 'x', NS.x).map((x) => integer(x.getAttribute('v'), 0)),
			prefix = integer(item.getAttribute('r'), 0);
		if (
			integer(item.getAttribute('i'), 0) !== measure ||
			!Number.isInteger(prefix) ||
			prefix < 0 ||
			prefix > Math.max(0, previous.length - 1)
		)
			throw new Error('Invalid page column prefix');
		if (grand) {
			if (item.getAttribute('t') !== 'grand' || prefix !== 0 || xs.length !== 1 || xs[0] !== 0)
				throw new Error('Unmodeled page grand item');
			previous = [];
			return;
		}
		const values =
			!columns && measures === 1 && xs.length === 0 && prefix === 0
				? [0]
				: [...previous.slice(0, prefix), ...xs];
		if (item.hasAttribute('t') || JSON.stringify(values) !== JSON.stringify(expected[i]))
			throw new Error('Unmodeled page column item');
		previous = values;
	});
}
