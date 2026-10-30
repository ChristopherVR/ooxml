import type { XmlElement } from '../../xml/index';
import { parseRange } from '../address';
import type { DifferentialStyle, SortBy, SortCondition, SortState } from '../model';
import { att, boolAttr, numAttr, xChildren } from './xml-util';

const SORT_BY = new Set<string>(['value', 'cellColor', 'fontColor', 'icon']);

/** Reads a `<sortState>`; `dxfs` resolves the differential format of a colour sort level. */
export function readSortState(
	node: XmlElement | undefined,
	dxfs?: readonly DifferentialStyle[],
): SortState | undefined {
	const range = parseRange(att(node, 'ref') ?? '');
	if (!node || !range) return undefined;
	const conditions: SortCondition[] = [];
	for (const child of xChildren(node, 'sortCondition')) {
		const ref = parseRange(att(child, 'ref') ?? '');
		if (!ref) continue;
		const condition: SortCondition = { ref };
		if (boolAttr(child, 'descending', false)) condition.descending = true;
		const by = att(child, 'sortBy');
		if (by && SORT_BY.has(by) && by !== 'value') condition.sortBy = by as SortBy;
		const customList = att(child, 'customList');
		if (customList) condition.customList = customList;
		const dxfId = numAttr(child, 'dxfId');
		if (dxfId !== undefined) {
			condition.dxfId = dxfId;
			const style = dxfs?.[dxfId];
			if (style) condition.style = structuredClone(style);
		}
		const iconSet = att(child, 'iconSet');
		if (iconSet) condition.iconSet = iconSet;
		const iconId = numAttr(child, 'iconId');
		if (iconId !== undefined) condition.iconId = iconId;
		conditions.push(condition);
	}
	return {
		range,
		...(boolAttr(node, 'columnSort', false) ? { columnSort: true } : {}),
		...(boolAttr(node, 'caseSensitive', false) ? { caseSensitive: true } : {}),
		conditions,
	};
}
