import {
	NS,
	buildXml,
	children,
	elements,
	first,
	parseXml,
	type XmlElement,
} from '../../xml/index';
import type { PivotTable, Slicer, SlicerCache } from '../model';
import { SLICER_DEFAULT_ROW_HEIGHT } from '../model-slicers';
import { XML_HEADER, attrs, escapeAttr } from '../write/xml-out';
import { pivotValueKey } from '../pivot/data';
import { TABLE_SLICER_URI } from './read-slicers';
import { hiddenItemKeys, itemsWithData } from '../pivot/item-filters';
import { pivotItems } from '../pivot/data';

export const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
export const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XR10 = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10';

/** `slicers` part: every attribute of a slicer comes from the model. */
export function slicersPartXml(slicers: readonly Slicer[]): string {
	const items = slicers
		.map((slicer) => {
			const values = attrs({
				name: slicer.name,
				'xr10:uid': slicer.uid,
				cache: slicer.cache,
				caption: slicer.caption,
				startItem: slicer.startItem || undefined,
				columnCount: slicer.columnCount && slicer.columnCount > 1 ? slicer.columnCount : undefined,
				showCaption: slicer.showCaption === false ? false : undefined,
				level: slicer.level || undefined,
				style: slicer.style,
				lockedPosition: slicer.lockedPosition || undefined,
				rowHeight: slicer.rowHeight ?? SLICER_DEFAULT_ROW_HEIGHT,
			});
			return slicer.extLstXml
				? `<slicer${values}>${slicer.extLstXml.replace(/^<extLst[^>]*>/, '<extLst>')}</slicer>`
				: `<slicer${values}/>`;
		})
		.join('');
	return `${XML_HEADER}<slicers xmlns="${X14}" xmlns:mc="${MC}" mc:Ignorable="x xr10" xmlns:x="${NS.x}" xmlns:xr10="${XR10}">${items}</slicers>`;
}

/** The pivot a pivot-slicer cache filters. */
export interface CacheFacts {
	pivot?: PivotTable;
	sheetId?: number;
}

function itemsXml(pivot: PivotTable, field: string): string {
	const all = pivotItems(pivot, field);
	const hidden = hiddenItemKeys(pivot, field);
	const withData = itemsWithData(pivot, field);
	const shared = new Map<string, number>();
	let index = 0;
	for (const record of pivot.records) {
		const key = pivotValueKey(record[pivot.fields.indexOf(field)] ?? null);
		if (!shared.has(key)) shared.set(key, index++);
	}
	const entries = all.map((value) => {
		const key = pivotValueKey(value);
		return `<i${attrs({ x: shared.get(key), s: hidden.has(key) ? undefined : true, nd: withData.has(key) ? undefined : true })}/>`;
	});
	return `<items count="${entries.length}">${entries.join('')}</items>`;
}

/** A slicer cache part: the source XML patched when there is one, else a fresh definition. */
export function slicerCacheXml(cache: SlicerCache, facts: CacheFacts): string {
	const source = cache.source;
	const doc = parseXml(
		cache.sourceXml ??
			`<slicerCacheDefinition xmlns="${X14}" xmlns:mc="${MC}" mc:Ignorable="x xr10" xmlns:x="${NS.x}" xmlns:xr10="${XR10}"/>`,
	);
	const root = doc.documentElement;
	root.setAttribute('name', cache.name);
	root.setAttribute('sourceName', cache.sourceName);
	if (cache.uid && !root.hasAttributeNS(XR10, 'uid'))
		root.setAttributeNS(XR10, 'xr10:uid', cache.uid);
	const make = (ns: string, name: string) => doc.createElementNS(ns, name);
	if (source.kind === 'pivot') {
		const pivot = facts.pivot;
		const list =
			children(root, 'pivotTables', X14)[0] ?? root.appendChild(make(X14, 'pivotTables'));
		while (list.firstChild) list.removeChild(list.firstChild);
		for (const ref of source.pivots) {
			const node = make(X14, 'pivotTable');
			node.setAttribute('tabId', String(facts.sheetId ?? ref.tabId));
			node.setAttribute('name', pivot?.name ?? ref.name);
			list.appendChild(node);
		}
		let data = children(root, 'data', X14)[0];
		if (!data) {
			data = make(X14, 'data');
			root.insertBefore(data, first(root, 'extLst', X14) ?? null);
		}
		let tabular = children(data, 'tabular', X14)[0];
		if (!tabular) tabular = data.appendChild(make(X14, 'tabular'));
		tabular.setAttribute('pivotCacheId', String(pivot?.slicerCacheId ?? source.pivotCacheId));
		if (cache.sortOrder === 'descending') tabular.setAttribute('sortOrder', 'descending');
		else tabular.removeAttribute('sortOrder');
		if (cache.hideNoData) tabular.setAttribute('showMissing', '0');
		else tabular.removeAttribute('showMissing');
		if (cache.crossFilter) tabular.setAttribute('crossFilter', cache.crossFilter);
		else tabular.removeAttribute('crossFilter');
		for (const old of children(tabular, 'items', X14)) tabular.removeChild(old);
		if (pivot) {
			const fragment = parseXml(`<r xmlns="${X14}">${itemsXml(pivot, cache.sourceName)}</r>`);
			tabular.appendChild(doc.importNode(elements(fragment.documentElement)[0]!, true));
		}
		return XML_HEADER + buildXml(root);
	}
	let list = first(root, 'extLst', X14);
	if (!list) {
		list = make(X14, 'extLst');
		root.appendChild(list);
	}
	let ext = elements(list).find((node) => node.getAttribute('uri') === TABLE_SLICER_URI);
	if (!ext) {
		ext = make(NS.x, 'x:ext');
		ext.setAttribute('uri', TABLE_SLICER_URI);
		ext.setAttribute('xmlns:x15', X15);
		list.appendChild(ext);
	}
	let table = first(ext, 'tableSlicerCache', X15);
	if (!table) table = ext.appendChild(make(X15, 'x15:tableSlicerCache'));
	table.setAttribute('tableId', String(source.tableId));
	table.setAttribute('column', String(source.column));
	if (cache.sortOrder === 'descending') table.setAttribute('sortOrder', 'descending');
	else table.removeAttribute('sortOrder');
	return XML_HEADER + buildXml(root);
}

/** The sheet-level extension entry listing slicer parts by relationship id. */
export function slicerListExt(uri: string, ids: readonly string[]): string {
	return `<ext uri="${uri}" xmlns:x14="${X14}"><x14:slicerList>${ids.map((id) => `<x14:slicer r:id="${escapeAttr(id)}"/>`).join('')}</x14:slicerList></ext>`;
}

export type { XmlElement };
