import { NS, children, first, type XmlElement } from '../../xml/index';
import type { CellValue, PivotItemName } from '../model';
import { att, decodeEscapes } from '../read/xml-util';

/**
 * The display names of items (`item/@n`): the cells show the name instead of the cached value.
 * Only data items of the cache's own fields carry one; the default subtotal item has none.
 */
export function readItemNames(
	tableFields: readonly XmlElement[],
	names: readonly string[],
	shared: readonly CellValue[][],
): PivotItemName[] {
	const out: PivotItemName[] = [];
	tableFields.forEach((field, index) => {
		const list = first(field, 'items', NS.x);
		if (!list || index >= names.length) return;
		for (const item of children(list, 'item', NS.x)) {
			const name = att(item, 'n');
			if (name === undefined) continue;
			const at = Number(att(item, 'x'));
			const value = shared[index]?.[at];
			if (att(item, 't') !== undefined || !Number.isInteger(at) || value === undefined)
				throw new Error('Unsupported pivot item name');
			out.push({ field: names[index]!, item: value, name: decodeEscapes(name) });
		}
	});
	return out;
}
