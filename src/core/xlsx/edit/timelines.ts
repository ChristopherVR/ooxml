import { createGridMetrics } from '../layout/metrics';
import { anchorToPixelBox, pixelBoxToAnchor } from '../layout/anchors';
import type { PivotTable, Timeline, TimelineCache } from '../model';
import { isoToSerial, serialToIso } from '../pivot/dates';
import { cloneBook, removeTimelines } from '../slicer/book';
import { slicerFrameXml } from '../slicer/frame';
import { boundsOfSerialRange, toIsoDate, type TimelineSelection } from '../slicer/timeline-periods';
import { timelineTarget } from '../slicer/timeline-view';
import { type EditContext, sheetAt } from './context';
import { writePivot } from './pivots';

const TIMELINE_BOX = { w: 480, h: 130 };

function cacheOf(ctx: EditContext, name: string): TimelineCache {
	const cache = ctx.workbook.slicers?.timelineCaches.find((entry) => entry.name === name);
	if (!cache) throw new Error('The timeline is not available.');
	return cache;
}

/**
 * Filters the PivotTable behind a timeline to a date range (`null` shows every date). One undo
 * step: the report is recomputed and the timeline cache records the range.
 */
export function setTimelineSelection(
	ctx: EditContext,
	cacheName: string,
	selection: TimelineSelection | null,
): void {
	const cache = cacheOf(ctx, cacheName);
	const target = timelineTarget(ctx.workbook, cache);
	if (!target) throw new Error('The timeline has no data source.');
	const { pivot } = target;
	const next: PivotTable = { ...pivot };
	if (selection) {
		const start = isoToSerial(selection.start, ctx.workbook.date1904);
		const end = isoToSerial(selection.end, ctx.workbook.date1904);
		if (start === undefined || end === undefined || start > end)
			throw new Error('Choose a valid timeline date range.');
		next.dateFilter = { field: cache.sourceName, start, end };
	} else delete next.dateFilter;
	writePivot(
		ctx,
		target.sheetIndex,
		next,
		pivot,
		selection ? 'Filter timeline' : 'Clear timeline filter',
		() => {
			const book = cloneBook(ctx.workbook);
			const entry = book.timelineCaches.find((item) => item.name === cacheName)!;
			if (selection) {
				entry.selection = { ...selection };
				entry.filterType = 'dateBetween';
			} else {
				delete entry.selection;
				entry.filterType = 'unknown';
			}
			book.dirty = true;
			ctx.workbook.slicers = book;
		},
	);
}

export interface TimelinePatch {
	caption?: string;
	level?: number;
	showHeader?: boolean;
	showSelectionLabel?: boolean;
	showTimeLevel?: boolean;
	showHorizontalScrollbar?: boolean;
	style?: string;
}

/** Changes a timeline's caption, time level or visible parts (one undo step). */
export function updateTimeline(ctx: EditContext, name: string, patch: TimelinePatch): void {
	if (!ctx.workbook.slicers?.timelines.some((entry) => entry.name === name))
		throw new Error('The timeline is not available.');
	if (
		(patch.level !== undefined && ![0, 1, 2, 3].includes(patch.level)) ||
		(patch.caption !== undefined && patch.caption.length > 255)
	)
		throw new Error('Choose a valid timeline setting.');
	ctx.run(
		'Format timeline',
		'view',
		[{ kind: 'meta' }],
		() => {
			const book = cloneBook(ctx.workbook);
			const entry = book.timelines.find((item) => item.name === name)!;
			for (const key of [
				'caption',
				'showHeader',
				'showSelectionLabel',
				'showTimeLevel',
				'showHorizontalScrollbar',
				'style',
			] as const)
				if (patch[key] !== undefined)
					(entry as unknown as Record<string, unknown>)[key] = patch[key];
			if (patch.level !== undefined) {
				entry.level = patch.level;
				entry.selectionLevel = patch.level;
				delete entry.scrollPosition;
			}
			book.dirty = true;
			ctx.workbook.slicers = book;
		},
		{ structural: false },
	);
}

/** Inserts a timeline for a date field of a PivotTable on sheet `s` (one undo step). */
export function insertTimeline(
	ctx: EditContext,
	s: number,
	ref: { pivot: string; field: string },
): string {
	const { workbook } = ctx;
	const sheet = sheetAt(workbook, s);
	const pivot = sheet.pivotTables?.find((entry) => entry.name === ref.pivot);
	if (!pivot || !pivot.dateFields?.includes(ref.field))
		throw new Error('Choose a PivotTable with a date field.');
	// The earliest and latest dates in a loop: a spread passes one argument per record and overflows.
	const column = pivot.fields.indexOf(ref.field);
	let earliest = Infinity;
	let latest = -Infinity;
	for (const record of pivot.records) {
		const value = record[column];
		if (typeof value !== 'number') continue;
		earliest = Math.min(earliest, value);
		latest = Math.max(latest, value);
	}
	if (earliest > latest) throw new Error('Choose a PivotTable with a date field.');
	const taken = new Set([
		...(workbook.slicers?.slicers ?? []).map((entry) => entry.name.toLowerCase()),
		...(workbook.slicers?.timelines ?? []).map((entry) => entry.name.toLowerCase()),
		...(workbook.slicers?.caches ?? []).map((entry) => entry.name.toLowerCase()),
		...(workbook.slicers?.timelineCaches ?? []).map((entry) => entry.name.toLowerCase()),
		...workbook.definedNames.map((entry) => entry.name.toLowerCase()),
	]);
	const unique = (base: string): string => {
		if (!taken.has(base.toLowerCase())) return base;
		for (let n = 1; ; n++) if (!taken.has(`${base} ${n}`.toLowerCase())) return `${base} ${n}`;
	};
	const name = unique(ref.field);
	const cacheName = unique(`NativeTimeline_${ref.field.replace(/[^\p{L}\p{N}_]/gu, '_')}`).replace(
		/ /g,
		'_',
	);
	const bounds = boundsOfSerialRange(
		serialToIso(earliest, workbook.date1904),
		serialToIso(latest, workbook.date1904),
	);
	const metrics = createGridMetrics(sheet, { zoom: 100 });
	let right = 0;
	for (const drawing of sheet.drawings)
		if (drawing.kind === 'slicer' || drawing.kind === 'timeline') {
			const box = anchorToPixelBox(sheet, drawing.anchor, metrics);
			right = Math.max(right, box.x + box.w);
		}
	const box = {
		x: Math.max(metrics.colLeft(pivot.location.end.col + 2), right + 10),
		y: metrics.rowTop(pivot.location.start.row),
		...TIMELINE_BOX,
	};
	const anchor = pixelBoxToAnchor(sheet, box, metrics, 'twoCell');
	const id =
		Math.max(
			1,
			...sheet.drawings.flatMap((d) =>
				'sourceXml' in d && d.sourceXml
					? [...d.sourceXml.matchAll(/<[\w:]*cNvPr\b[^>]*>/g)].flatMap((m) => {
							const id = /\sid="(\d+)"/.exec(m[0])?.[1];
							return id ? [Number(id)] : [];
						})
					: [],
			),
		) + 1;
	ctx.run(
		'Insert timeline',
		'annotations',
		[{ kind: 'sheet', sheet: s }, { kind: 'meta' }],
		() => {
			const book = cloneBook(workbook);
			pivot.slicerCacheId ??=
				Math.max(
					1_000_000,
					...workbook.sheets.flatMap((sh) =>
						(sh.pivotTables ?? []).map((p) => p.slicerCacheId ?? 0),
					),
				) + 1;
			pivot.dirty = true;
			const cache: TimelineCache = {
				name: cacheName,
				sourceName: ref.field,
				pivotCacheId: pivot.slicerCacheId,
				pivots: [{ tabId: sheet.sheetId, name: pivot.name }],
				bounds,
				filterType: 'unknown',
				minimalRefreshVersion: 6,
				lastRefreshVersion: 6,
			};
			const timeline: Timeline = {
				name,
				cache: cacheName,
				caption: ref.field,
				level: 2,
				selectionLevel: 2,
				scrollPosition: toIsoDate(bounds.start),
				sheet: s,
			};
			book.timelineCaches.push(cache);
			book.timelines.push(timeline);
			book.dirty = true;
			workbook.slicers = book;
			workbook.definedNames.push({ name: cacheName, formula: '#N/A' });
			sheet.drawings.push({
				kind: 'timeline',
				anchor,
				name,
				sourceXml: slicerFrameXml('timeline', name, id, anchor, {
					x: box.x * 9525,
					y: box.y * 9525,
					cx: box.w * 9525,
					cy: box.h * 9525,
				}),
			});
		},
		{ sheet: s },
	);
	return name;
}

/** Deletes a timeline and its cache; the report shows every date again (one undo step). */
export function deleteTimeline(ctx: EditContext, name: string): void {
	const timeline = ctx.workbook.slicers?.timelines.find((entry) => entry.name === name);
	if (!timeline) throw new Error('The timeline is not available.');
	const cache = ctx.workbook.slicers?.timelineCaches.find((entry) => entry.name === timeline.cache);
	const target = cache && timelineTarget(ctx.workbook, cache);
	if (cache?.selection && target) {
		// Clearing the filter first keeps the report consistent with the removed timeline.
		setTimelineSelection(ctx, cache.name, null);
	}
	ctx.run(
		'Delete timeline',
		'annotations',
		[{ kind: 'workbook' }],
		() => removeTimelines(ctx.workbook, (entry) => entry.name === name),
		{},
	);
}
