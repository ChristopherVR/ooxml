import type { CellValue, PivotItemName, PivotTable } from '../model';
import { type CaptionHolder, pivotCaptions } from './captions';
import { pivotItemLabel, pivotValueKey } from './data';

const keyOf = (field: string, item: CellValue) => `${field}\u0000${pivotValueKey(item)}`;

/**
 * The text a field's item shows in the cells: the display name the file gave it (`item/@n`),
 * else the value, a blank item as the report's blank caption.
 */
export function pivotItemLabeler(
	pivot: CaptionHolder & { itemNames?: readonly PivotItemName[] | undefined },
): (field: string, value: CellValue) => CellValue {
	const blank = pivotCaptions(pivot).blank;
	const names = pivot.itemNames?.length
		? new Map(pivot.itemNames.map((entry) => [keyOf(entry.field, entry.item), entry.name]))
		: undefined;
	return (field, value) =>
		names?.get(keyOf(field, value)) ??
		(value === null || value === '' ? blank : pivotItemLabel(value));
}

/** The display name of one item, when the report names it. */
export const pivotItemName = (
	pivot: Pick<PivotTable, 'itemNames'>,
	field: string,
	item: CellValue,
): string | undefined =>
	pivot.itemNames?.find(
		(entry) => entry.field === field && pivotValueKey(entry.item) === pivotValueKey(item),
	)?.name;
