import { RELATIONSHIP_TYPES } from '../../opc/index';
import { elements, parseXml } from '../../xml/index';
import type {
	SlicerBook,
	SlicerCache,
	Slicer,
	Timeline,
	TimelineCache,
	UnsupportedDrawingObject,
	Workbook,
} from '../model';
import type { SourceIndex } from '../read/package';
import { parseSlicer, parseSlicerCache } from './read-slicers';
import { parseTimeline, parseTimelineCache } from './read-timelines';
import { readTableFilteredRows } from './table-filter';

export const SLICER_RELATIONSHIPS = {
	slicer: 'http://schemas.microsoft.com/office/2007/relationships/slicer',
	slicerCache: 'http://schemas.microsoft.com/office/2007/relationships/slicerCache',
	timeline: 'http://schemas.microsoft.com/office/2011/relationships/timeline',
	timelineCache: 'http://schemas.microsoft.com/office/2011/relationships/timelineCache',
} as const;

/** `tabId:name` keys of the PivotTables a slicer or timeline cache connects to. */
export const pivotKey = (tabId: number, name: string): string => `${tabId}:${name}`;

/** What reading found, for the PivotTable reader: which pivots slicers may stay attached to. */
export interface SlicerRefs {
	/** Pivots connected only to caches this model understands. */
	supported: Set<string>;
	/** Pivots some unmodelled slicer, timeline or cache points at. */
	blocked: Set<string>;
}

function rawPivotRefs(xml: string | undefined): string[] {
	if (!xml) return [];
	try {
		const root = parseXml(xml).documentElement;
		const list = elements(root).find((node) => node.localName === 'pivotTables');
		return list
			? elements(list).map((node) =>
					pivotKey(Number(node.getAttribute('tabId')), node.getAttribute('name') ?? ''),
				)
			: [];
	} catch {
		return [];
	}
}

/** Reads the slicer and timeline parts into `workbook.slicers`; unreadable ones stay verbatim. */
export function readSlicers(workbook: Workbook, source: SourceIndex): SlicerRefs {
	const refs: SlicerRefs = { supported: new Set(), blocked: new Set() };
	const bookPart = source.workbookPart();
	if (!bookPart) return refs;
	const book: SlicerBook = { caches: [], slicers: [], timelineCaches: [], timelines: [] };
	const unmodeledCaches = new Set<string>();
	for (const rel of source.rels(bookPart).values()) {
		const target = source.target(bookPart, rel);
		const xml = target ? source.text(target) : undefined;
		if (!target || !xml) continue;
		if (rel.type === SLICER_RELATIONSHIPS.slicerCache) {
			const cache = parseSlicerCache(xml, target);
			if (cache) {
				book.caches.push(cache);
				continue;
			}
			unmodeledCaches.add(target);
		} else if (rel.type === SLICER_RELATIONSHIPS.timelineCache) {
			const cache = parseTimelineCache(xml, target);
			if (cache) {
				book.timelineCaches.push(cache);
				continue;
			}
			unmodeledCaches.add(target);
		} else continue;
		for (const ref of rawPivotRefs(xml)) refs.blocked.add(ref);
	}
	const unmodeledParts = new Set<string>();
	workbook.sheets.forEach((sheet, index) => {
		if (!sheet.partName) return;
		for (const rel of source.rels(sheet.partName).values()) {
			const target = source.target(sheet.partName, rel);
			const xml = target ? source.text(target) : undefined;
			if (!target || !xml) continue;
			const isSlicers = rel.type === SLICER_RELATIONSHIPS.slicer;
			if (!isSlicers && rel.type !== SLICER_RELATIONSHIPS.timeline) continue;
			try {
				const entries = elements(parseXml(xml).documentElement);
				for (const node of entries) {
					if (isSlicers) {
						const slicer = parseSlicer(node, index, target);
						if (!slicer) throw new Error('Unreadable slicer');
						book.slicers.push(slicer);
					} else {
						const timeline = parseTimeline(node, index, target);
						if (!timeline) throw new Error('Unreadable timeline');
						book.timelines.push(timeline);
					}
				}
			} catch {
				unmodeledParts.add(target);
			}
		}
	});
	// A part is modelled whole or kept whole: drop what it listed when any entry was unreadable.
	book.slicers = book.slicers.filter((s) => !unmodeledParts.has(s.partName ?? ''));
	book.timelines = book.timelines.filter((t) => !unmodeledParts.has(t.partName ?? ''));
	if (
		book.caches.length ||
		book.slicers.length ||
		book.timelineCaches.length ||
		book.timelines.length
	)
		workbook.slicers = book;
	readTableFilteredRows(workbook);
	for (const cache of [...book.caches, ...book.timelineCaches])
		if (cache.partName && !unmodeledCaches.has(cache.partName))
			for (const pivot of 'source' in cache
				? cache.source.kind === 'pivot'
					? cache.source.pivots
					: []
				: cache.pivots)
				refs.supported.add(pivotKey(pivot.tabId, pivot.name));
	return refs;
}

/** Restores a frame whose slicer cannot be edited to an honest "not shown" placeholder. */
function placeholder(
	frame: { anchor: UnsupportedDrawingObject['anchor']; sourceXml: string },
	kind: string,
): UnsupportedDrawingObject {
	return {
		kind: 'unsupported',
		anchor: frame.anchor,
		description: kind,
		sourceXml: frame.sourceXml,
	};
}

/**
 * After tables and PivotTables are read: keep only slicers and timelines whose target is an
 * editable table column or an admitted PivotTable; the rest are kept verbatim and drawn as
 * placeholders.
 */
export function settleSlicers(workbook: Workbook, warn: (message: string) => void): void {
	const book = workbook.slicers;
	if (!book) return;
	const pivotAt = (tabId: number, name: string) => {
		const sheet = workbook.sheets.find((entry) => entry.sheetId === tabId);
		const pivot = sheet?.pivotTables?.find((entry) => entry.name === name);
		return sheet && pivot ? { sheet, pivot } : undefined;
	};
	const good = (cache: SlicerCache): boolean => {
		if (cache.source.kind === 'table') {
			const { tableId, column } = cache.source;
			for (const sheet of workbook.sheets) {
				const table = sheet.tables.find((entry) => entry.id === tableId);
				if (table)
					return (
						table.headerRow &&
						!table.filterNotModeled &&
						table.columns.some((entry) => entry.sourceId === column)
					);
			}
			return false;
		}
		const { pivots, pivotCacheId } = cache.source;
		const found = pivots.map((p) => pivotAt(p.tabId, p.name));
		if (found.length !== 1 || !found[0]) return false;
		const { pivot } = found[0];
		if (pivot.slicerCacheId !== pivotCacheId || !pivot.fields.includes(cache.sourceName))
			return false;
		return pivot.pageFilter?.field !== cache.sourceName;
	};
	const goodTimeline = (cache: TimelineCache): boolean => {
		const found = cache.pivots.map((p) => pivotAt(p.tabId, p.name));
		return found.length === 1 && !!found[0] && found[0].pivot.slicerCacheId === cache.pivotCacheId;
	};
	const caches = book.caches.filter(good);
	const timelineCaches = book.timelineCaches.filter(goodTimeline);
	const slicers: Slicer[] = [];
	const timelines: Timeline[] = [];
	const partOk = new Map<string, boolean>();
	for (const slicer of book.slicers) {
		const ok = caches.some((c) => c.name === slicer.cache);
		partOk.set(slicer.partName ?? '', (partOk.get(slicer.partName ?? '') ?? true) && ok);
	}
	for (const timeline of book.timelines) {
		const ok = timelineCaches.some((c) => c.name === timeline.cache);
		partOk.set(timeline.partName ?? '', (partOk.get(timeline.partName ?? '') ?? true) && ok);
	}
	for (const slicer of book.slicers) if (partOk.get(slicer.partName ?? '')) slicers.push(slicer);
	for (const timeline of book.timelines)
		if (partOk.get(timeline.partName ?? '')) timelines.push(timeline);
	const usedCaches = new Set(slicers.map((s) => s.cache));
	const usedTimelineCaches = new Set(timelines.map((t) => t.cache));
	let dropped = false;
	for (const sheet of workbook.sheets)
		sheet.drawings = sheet.drawings.map((drawing) => {
			if (drawing.kind !== 'slicer' && drawing.kind !== 'timeline') return drawing;
			const modeled =
				drawing.kind === 'slicer'
					? slicers.some((s) => s.name === drawing.name)
					: timelines.some((t) => t.name === drawing.name);
			if (modeled) return drawing;
			dropped = true;
			return placeholder(drawing, drawing.kind);
		});
	book.caches = caches.filter((cache) => usedCaches.has(cache.name));
	book.timelineCaches = timelineCaches.filter((cache) => usedTimelineCaches.has(cache.name));
	book.slicers = slicers;
	book.timelines = timelines;
	if (dropped)
		warn(
			'Some slicers or timelines are kept but not shown: their table, PivotTable or data source is not supported.',
		);
	if (!book.caches.length && !book.timelineCaches.length) delete workbook.slicers;
}
