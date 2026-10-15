import { NS, children, elements, first, parseXml, type XmlElement } from '../../xml/index';
import type { Slicer, SlicerCache } from '../model';
import { att, boolAttr, numAttr, selfContainedXml } from '../read/xml-util';

export const TABLE_SLICER_URI = '{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}';
const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';

const byLocal = (parent: XmlElement | null | undefined, local: string): XmlElement | undefined =>
	parent ? elements(parent).find((node) => node.localName === local) : undefined;

/** One `slicer` element; `sheet` is the sheet whose part lists it. */
export function parseSlicer(node: XmlElement, sheet: number, partName: string): Slicer | undefined {
	const name = att(node, 'name');
	const cache = att(node, 'cache');
	if (!name || !cache) return undefined;
	const extLst = byLocal(node, 'extLst');
	const slicer: Slicer = { name, cache, sheet, partName };
	const uid = node.getAttributeNS(
		'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',
		'uid',
	);
	if (uid) slicer.uid = uid;
	const caption = att(node, 'caption');
	if (caption !== undefined) slicer.caption = caption;
	const style = att(node, 'style');
	if (style) slicer.style = style;
	const columns = numAttr(node, 'columnCount');
	if (columns !== undefined) slicer.columnCount = columns;
	const rowHeight = numAttr(node, 'rowHeight');
	if (rowHeight !== undefined) slicer.rowHeight = rowHeight;
	const showCaption = boolAttr(node, 'showCaption');
	if (showCaption !== undefined) slicer.showCaption = showCaption;
	const start = numAttr(node, 'startItem');
	if (start !== undefined) slicer.startItem = start;
	const locked = boolAttr(node, 'lockedPosition');
	if (locked !== undefined) slicer.lockedPosition = locked;
	const level = numAttr(node, 'level');
	if (level !== undefined) slicer.level = level;
	if (extLst) slicer.extLstXml = selfContainedXml(extLst);
	return slicer;
}

/**
 * Reads a slicer cache part. `undefined` means a source this model does not edit (OLAP, a
 * hierarchy, an unknown source); the caller then keeps the slicers and their parts verbatim.
 */
export function parseSlicerCache(xml: string, partName: string): SlicerCache | undefined {
	let root: XmlElement;
	try {
		root = parseXml(xml, { label: 'XLSX slicer cache' }).documentElement;
	} catch {
		return undefined;
	}
	const name = att(root, 'name');
	const sourceName = att(root, 'sourceName');
	if (root.localName !== 'slicerCacheDefinition' || !name || !sourceName) return undefined;
	if (att(root, 'sourceName')?.startsWith('[')) return undefined;
	const uid = root.getAttributeNS(
		'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',
		'uid',
	);
	const common = { name, sourceName, partName, sourceXml: xml, ...(uid ? { uid } : {}) };
	const data = byLocal(root, 'data');
	const tabular = byLocal(data, 'tabular');
	if (tabular) {
		if (byLocal(data, 'olap')) return undefined;
		const pivots = children(byLocal(root, 'pivotTables') ?? root, 'pivotTable', NS.x14).map(
			(node) => ({ tabId: numAttr(node, 'tabId') ?? -1, name: att(node, 'name') ?? '' }),
		);
		const pivotCacheId = numAttr(tabular, 'pivotCacheId');
		if (!pivots.length || pivotCacheId === undefined || pivots.some((p) => p.tabId < 0 || !p.name))
			return undefined;
		const cache: SlicerCache = {
			...common,
			source: { kind: 'pivot', pivotCacheId, pivots },
		};
		const order = att(tabular, 'sortOrder');
		if (order === 'descending') cache.sortOrder = 'descending';
		if (att(tabular, 'showMissing') === '0') cache.hideNoData = true;
		const cross = att(tabular, 'crossFilter');
		if (cross === 'none' || cross === 'showItemsWithNoData' || cross === 'showItemsWithDataAtTop')
			cache.crossFilter = cross;
		return cache;
	}
	if (data) return undefined;
	const list = byLocal(root, 'extLst');
	const ext = list
		? elements(list).find((node) => node.getAttribute('uri') === TABLE_SLICER_URI)
		: undefined;
	const table = ext ? first(ext, 'tableSlicerCache', X15) : undefined;
	const tableId = numAttr(table, 'tableId');
	const column = numAttr(table, 'column');
	if (!table || tableId === undefined || column === undefined) return undefined;
	const cache: SlicerCache = { ...common, source: { kind: 'table', tableId, column } };
	if (att(table, 'sortOrder') === 'descending') cache.sortOrder = 'descending';
	if (elements(root).some((node) => node.localName === 'pivotTables')) return undefined;
	return cache;
}
