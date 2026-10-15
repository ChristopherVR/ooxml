import { children, elements, parseXml, type XmlElement } from '../../xml/index';
import type { Timeline, TimelineCache } from '../model';
import { att, boolAttr, numAttr, selfContainedXml } from '../read/xml-util';

const REV10 = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10';
const byLocal = (parent: XmlElement | null | undefined, local: string): XmlElement | undefined =>
	parent ? elements(parent).find((node) => node.localName === local) : undefined;

export function parseTimeline(
	node: XmlElement,
	sheet: number,
	partName: string,
): Timeline | undefined {
	const name = att(node, 'name');
	const cache = att(node, 'cache');
	if (!name || !cache) return undefined;
	const timeline: Timeline = {
		name,
		cache,
		sheet,
		partName,
		level: numAttr(node, 'level') ?? 2,
		selectionLevel: numAttr(node, 'selectionLevel') ?? numAttr(node, 'level') ?? 2,
	};
	const uid = node.getAttributeNS(REV10, 'uid');
	if (uid) timeline.uid = uid;
	for (const key of ['caption', 'scrollPosition', 'style'] as const) {
		const value = att(node, key);
		if (value !== undefined) timeline[key] = value;
	}
	for (const key of [
		'showHeader',
		'showSelectionLabel',
		'showTimeLevel',
		'showHorizontalScrollbar',
	] as const) {
		const value = boolAttr(node, key);
		if (value !== undefined) timeline[key] = value;
	}
	const extLst = byLocal(node, 'extLst');
	if (extLst) timeline.extLstXml = selfContainedXml(extLst);
	return timeline;
}

export function parseTimelineCache(xml: string, partName: string): TimelineCache | undefined {
	let root: XmlElement;
	try {
		root = parseXml(xml, { label: 'XLSX timeline cache' }).documentElement;
	} catch {
		return undefined;
	}
	const name = att(root, 'name');
	const sourceName = att(root, 'sourceName');
	const state = byLocal(root, 'state');
	const bounds = byLocal(state, 'bounds');
	const pivotCacheId = numAttr(state, 'pivotCacheId');
	const pivots = children(
		byLocal(root, 'pivotTables') ?? root,
		'pivotTable',
		root.namespaceURI ?? '',
	).map((node) => ({ tabId: numAttr(node, 'tabId') ?? -1, name: att(node, 'name') ?? '' }));
	const start = att(bounds, 'startDate');
	const end = att(bounds, 'endDate');
	if (
		root.localName !== 'timelineCacheDefinition' ||
		!name ||
		!sourceName ||
		!state ||
		pivotCacheId === undefined ||
		!pivots.length ||
		pivots.some((p) => p.tabId < 0 || !p.name) ||
		!start ||
		!end
	)
		return undefined;
	const cache: TimelineCache = {
		name,
		sourceName,
		pivotCacheId,
		pivots,
		bounds: { start, end },
		partName,
	};
	const uid = root.getAttributeNS(REV10, 'uid');
	if (uid) cache.uid = uid;
	const selection = byLocal(state, 'selection');
	const from = att(selection, 'startDate');
	const to = att(selection, 'endDate');
	if (from && to) cache.selection = { start: from, end: to };
	const filter = att(state, 'filterType');
	if (filter) cache.filterType = filter;
	const minimal = numAttr(state, 'minimalRefreshVersion');
	if (minimal !== undefined) cache.minimalRefreshVersion = minimal;
	const last = numAttr(state, 'lastRefreshVersion');
	if (last !== undefined) cache.lastRefreshVersion = last;
	return cache;
}
