import type { SlicerItem } from './items';

export type SlicerClick = 'single' | 'toggle' | 'range';

/**
 * The keys that stay selected after a click, or `null` for "everything" (no filter). A plain click
 * selects one item; Ctrl+click (or Multi-Select mode) toggles; Shift+click extends from the
 * anchor. Emptying the selection returns to everything, as Excel does.
 */
export function selectionAfterClick(
	items: readonly SlicerItem[],
	clicked: string,
	mode: SlicerClick,
	anchor?: string,
): string[] | null {
	const target = items.find((item) => item.key === clicked);
	if (!target) return items.every((item) => item.selected) ? null : selectedKeys(items);
	const everything = items.every((item) => item.selected);
	let next: string[];
	if (mode === 'toggle') {
		const current = new Set(selectedKeys(items));
		if (current.has(clicked)) current.delete(clicked);
		else current.add(clicked);
		next = items.map((item) => item.key).filter((key) => current.has(key));
	} else if (mode === 'range' && anchor !== undefined) {
		const from = items.findIndex((item) => item.key === anchor);
		const to = items.findIndex((item) => item.key === clicked);
		if (from < 0) next = [clicked];
		else next = items.slice(Math.min(from, to), Math.max(from, to) + 1).map((item) => item.key);
	} else if (!everything && selectedKeys(items).length === 1 && target.selected) next = [];
	else next = [clicked];
	return next.length === 0 || next.length === items.length ? null : next;
}

export const selectedKeys = (items: readonly SlicerItem[]): string[] =>
	items.filter((item) => item.selected).map((item) => item.key);
