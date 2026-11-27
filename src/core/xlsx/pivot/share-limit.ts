// How large a PivotTable a collaboration room carries. A shared PivotTable travels whole (its
// definition with the cached source fields and records), so one with a large cache is kept on
// the peer that has it: it is saved and edited there as before, but only while not shared.
import type { PivotTable } from '../model-pivots';

/** The largest shared PivotTable, measured as the JSON characters of its definition and cache. */
export const MAX_SHARED_PIVOT_CHARS = 4 * 1024 * 1024;

/** The shared form of a PivotTable: everything but what names the local package's parts. */
export function sharedPivotForm(pivot: PivotTable): Record<string, unknown> {
	const { partName: _part, cacheId: _cache, ...rest } = pivot;
	return JSON.parse(JSON.stringify(rest)) as Record<string, unknown>;
}

const sizes = new WeakMap<PivotTable, number>();

/** The JSON characters the shared form of `pivot` takes (cached per PivotTable object). */
export function sharedPivotSize(pivot: PivotTable): number {
	let size = sizes.get(pivot);
	if (size === undefined) {
		const { partName: _part, cacheId: _cache, ...rest } = pivot;
		size = JSON.stringify(rest).length;
		sizes.set(pivot, size);
	}
	return size;
}

/** Whether `pivot` is small enough to be shared in a collaboration room. */
export const pivotShareable = (pivot: PivotTable): boolean =>
	sharedPivotSize(pivot) <= MAX_SHARED_PIVOT_CHARS;
