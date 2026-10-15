import type { Slicer, SlicerBook, Timeline, Workbook } from '../model';

/** A private copy of the slicer book (edits replace it so history can restore the old one). */
export function cloneBook(workbook: Workbook): SlicerBook {
	const book = workbook.slicers;
	return book
		? structuredClone(book)
		: { caches: [], slicers: [], timelineCaches: [], timelines: [] };
}

/** Replaces the slicer book with a copy flagged for rewriting on save. */
export function markDirty(workbook: Workbook): void {
	if (!workbook.slicers) return;
	const book = structuredClone(workbook.slicers);
	book.dirty = true;
	workbook.slicers = book;
}

/** Whether a PivotTable has slicers or timelines attached. */
export function pivotHasSlicers(workbook: Workbook, tabId: number, name: string): boolean {
	const book = workbook.slicers;
	const same = (pivots: readonly { tabId: number; name: string }[]) =>
		pivots.some((pivot) => pivot.tabId === tabId && pivot.name === name);
	return (
		!!book?.caches.some((c) => c.source.kind === 'pivot' && same(c.source.pivots)) ||
		!!book?.timelineCaches.some((c) => same(c.pivots))
	);
}

/** Removes the timelines `doomed` selects, their frames and the caches nothing else uses. */
export function removeTimelines(workbook: Workbook, doomed: (timeline: Timeline) => boolean): void {
	const current = workbook.slicers;
	if (!current) return;
	const book = structuredClone(current);
	const removed = book.timelines.filter(doomed);
	if (!removed.length) return;
	const gone = new Set(removed.map((timeline) => timeline.name));
	book.timelines = book.timelines.filter((timeline) => !gone.has(timeline.name));
	const parts = new Set(book.removedParts ?? []);
	for (const part of new Set(removed.map((timeline) => timeline.partName)))
		if (part && !book.timelines.some((timeline) => timeline.partName === part)) parts.add(part);
	const used = new Set(book.timelines.map((timeline) => timeline.cache));
	const dropped = book.timelineCaches.filter((cache) => !used.has(cache.name));
	book.timelineCaches = book.timelineCaches.filter((cache) => used.has(cache.name));
	for (const cache of dropped) if (cache.partName) parts.add(cache.partName);
	workbook.definedNames = workbook.definedNames.filter(
		(name) => !dropped.some((cache) => cache.name === name.name && name.formula === '#N/A'),
	);
	if (parts.size) book.removedParts = [...parts];
	book.dirty = true;
	for (const sheet of workbook.sheets)
		sheet.drawings = sheet.drawings.filter((d) => !(d.kind === 'timeline' && gone.has(d.name)));
	workbook.slicers = book;
}

/** Whether the timeline cache filters the given PivotTable. */
export function timelineOnPivot(
	workbook: Workbook,
	cacheName: string,
	tabId: number,
	name: string,
): boolean {
	const cache = workbook.slicers?.timelineCaches.find((entry) => entry.name === cacheName);
	return !!cache?.pivots.some((pivot) => pivot.tabId === tabId && pivot.name === name);
}

/** Whether the slicer cache filters the given PivotTable. */
export function slicerOnPivot(
	workbook: Workbook,
	cacheName: string,
	tabId: number,
	name: string,
): boolean {
	const cache = workbook.slicers?.caches.find((entry) => entry.name === cacheName);
	return (
		cache?.source.kind === 'pivot' &&
		cache.source.pivots.some((pivot) => pivot.tabId === tabId && pivot.name === name)
	);
}

/** Removes the slicers `doomed` selects, their frames and the caches nothing else uses. */
export function removeSlicers(workbook: Workbook, doomed: (slicer: Slicer) => boolean): void {
	const current = workbook.slicers;
	if (!current) return;
	const book = structuredClone(current);
	const removed = book.slicers.filter(doomed);
	if (!removed.length) return;
	const gone = new Set(removed.map((slicer) => slicer.name));
	book.slicers = book.slicers.filter((slicer) => !gone.has(slicer.name));
	const parts = new Set(book.removedParts ?? []);
	for (const part of new Set(removed.map((slicer) => slicer.partName)))
		if (part && !book.slicers.some((slicer) => slicer.partName === part)) parts.add(part);
	const used = new Set(book.slicers.map((slicer) => slicer.cache));
	const dropped = book.caches.filter((cache) => !used.has(cache.name));
	book.caches = book.caches.filter((cache) => used.has(cache.name));
	for (const cache of dropped) if (cache.partName) parts.add(cache.partName);
	workbook.definedNames = workbook.definedNames.filter(
		(name) => !dropped.some((cache) => cache.name === name.name && name.formula === '#N/A'),
	);
	if (parts.size) book.removedParts = [...parts];
	book.dirty = true;
	for (const sheet of workbook.sheets)
		sheet.drawings = sheet.drawings.filter((d) => !(d.kind === 'slicer' && gone.has(d.name)));
	workbook.slicers = book;
}
