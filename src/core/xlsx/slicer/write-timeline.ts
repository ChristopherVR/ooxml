import { buildXml, elements, parseXml } from '../../xml/index';
import type { Timeline, TimelineCache } from '../model';
import { XML_HEADER, attrs } from '../write/xml-out';

export const X15_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const XR10 = 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10';
const X = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** `timelines` part: every attribute of a timeline comes from the model. */
export function timelinesPartXml(timelines: readonly Timeline[]): string {
	const items = timelines
		.map((timeline) => {
			const values = attrs({
				name: timeline.name,
				'xr10:uid': timeline.uid,
				cache: timeline.cache,
				caption: timeline.caption,
				showHeader: timeline.showHeader === false ? false : undefined,
				showSelectionLabel: timeline.showSelectionLabel === false ? false : undefined,
				showTimeLevel: timeline.showTimeLevel === false ? false : undefined,
				showHorizontalScrollbar: timeline.showHorizontalScrollbar === false ? false : undefined,
				level: timeline.level,
				selectionLevel: timeline.selectionLevel,
				scrollPosition: timeline.scrollPosition,
				style: timeline.style,
			});
			return timeline.extLstXml
				? `<timeline${values}>${timeline.extLstXml.replace(/^<extLst[^>]*>/, '<extLst>')}</timeline>`
				: `<timeline${values}/>`;
		})
		.join('');
	return `${XML_HEADER}<timelines xmlns="${X15_NS}" xmlns:mc="${MC}" mc:Ignorable="x xr10" xmlns:x="${X}" xmlns:xr10="${XR10}">${items}</timelines>`;
}

/** A timeline cache part: the source XML patched when there is one, else a fresh definition. */
export function timelineCacheXml(cache: TimelineCache, pivotName: string, tabId: number): string {
	const doc = parseXml(
		cache.sourceXml ??
			`<timelineCacheDefinition xmlns="${X15_NS}" xmlns:x15="${X15_NS}" xmlns:mc="${MC}" mc:Ignorable="xr10" xmlns:xr10="${XR10}"><pivotTables/><state/></timelineCacheDefinition>`,
	);
	const root = doc.documentElement;
	root.setAttribute('name', cache.name);
	root.setAttribute('sourceName', cache.sourceName);
	const make = (name: string) => doc.createElementNS(root.namespaceURI, name);
	let tables = elements(root).find((node) => node.localName === 'pivotTables');
	if (!tables) tables = root.insertBefore(make('pivotTables'), root.firstChild);
	while (tables.firstChild) tables.removeChild(tables.firstChild);
	for (const ref of cache.pivots) {
		const node = make('pivotTable');
		node.setAttribute('tabId', String(tabId || ref.tabId));
		node.setAttribute('name', pivotName || ref.name);
		tables.appendChild(node);
	}
	let state = elements(root).find((node) => node.localName === 'state');
	if (!state) state = root.appendChild(make('state'));
	state.setAttribute('minimalRefreshVersion', String(cache.minimalRefreshVersion ?? 6));
	state.setAttribute('lastRefreshVersion', String(cache.lastRefreshVersion ?? 6));
	state.setAttribute('pivotCacheId', String(cache.pivotCacheId));
	state.setAttribute('filterType', cache.filterType ?? 'unknown');
	for (const old of elements(state)) state.removeChild(old);
	if (cache.selection) {
		const selection = make('selection');
		selection.setAttribute('startDate', cache.selection.start);
		selection.setAttribute('endDate', cache.selection.end);
		state.appendChild(selection);
	}
	const bounds = make('bounds');
	bounds.setAttribute('startDate', cache.bounds.start);
	bounds.setAttribute('endDate', cache.bounds.end);
	state.appendChild(bounds);
	return XML_HEADER + buildXml(root);
}
