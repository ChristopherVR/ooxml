import { NS, children, elements, first, type XmlElement } from '../../xml/index';
import type { CellValue, PivotItemFilter } from '../model';
import { att } from '../read/xml-util';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

/** `x14:pivotCacheDefinition/@pivotCacheId`: the identity slicers find the cache by. */
export function readSlicerCacheId(cache: XmlElement): number | undefined {
	for (const extension of elements(first(cache, 'extLst', NS.x) ?? cache)) {
		const id = att(first(extension, 'pivotCacheDefinition', X14), 'pivotCacheId');
		if (id !== undefined && /^\d+$/.test(id)) return Number(id);
	}
	return undefined;
}

/** Hidden items (`item/@h`) of every field, as item filters on the cached values. */
export function readItemFilters(
	fields: XmlElement[],
	names: string[],
	shared: CellValue[][],
): PivotItemFilter[] | undefined {
	const filters: PivotItemFilter[] = [];
	fields.forEach((field, index) => {
		const list = first(field, 'items', NS.x);
		if (!list) return;
		const hidden: CellValue[] = [];
		for (const item of children(list, 'item', NS.x)) {
			if (att(item, 'h') !== '1') continue;
			const x = att(item, 'x');
			const value = x !== undefined && /^\d+$/.test(x) ? shared[index]?.[Number(x)] : undefined;
			if (value === undefined) throw new Error('Invalid hidden pivot item');
			hidden.push(value);
		}
		const name = names[index];
		if (hidden.length && name) filters.push({ field: name, hidden });
	});
	return filters.length ? filters : undefined;
}
