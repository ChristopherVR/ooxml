import { NS, buildXml, elements, parseXml } from '../../xml/index';
import { formatRange } from '../address';
import type { PivotTable } from '../model';
import { XML_HEADER, attrs, encodeEscapes } from '../write/xml-out';
import { pivotGrid } from './aggregate';
import { pivotValueKey } from './data';
import type { PivotCacheData } from './cache-xml';

/** Writes the supported table structure while retaining unrelated native attributes/extensions. */
export function pivotTableXml(pivot: PivotTable, cacheId: number, cache: PivotCacheData): string {
	const grid = pivotGrid(pivot);
	const row = pivot.fields.indexOf(pivot.rowField);
	const col = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const value = pivot.fields.indexOf(pivot.valueField);
	const fields = pivot.fields
		.map((_field, index) => {
			const items = index === row ? grid.rowItems : index === col ? grid.columnItems : undefined;
			return `<pivotField${attrs({ axis: index === row ? 'axisRow' : index === col ? 'axisCol' : undefined, dataField: index === value ? true : undefined, compact: false, outline: false, showAll: false })}>${items ? `<items count="${items.length + 1}">${items.map((item) => `<item x="${cache.indices[index]!.get(pivotValueKey(item))}"/>`).join('')}<item t="default"/></items>` : ''}</pivotField>`;
		})
		.join('');
	const items = (count: number) =>
		Array.from({ length: count }, (_v, i) => `<i><x v="${i}"/></i>`).join('') +
		'<i t="grand"><x/></i>';
	const header = grid.firstDataRow - 1;
	let body = `<location${attrs({ ref: formatRange(pivot.location), firstHeaderRow: header, firstDataRow: grid.firstDataRow, firstDataCol: 1 })}/><pivotFields count="${pivot.fields.length}">${fields}</pivotFields><rowFields count="1"><field x="${row}"/></rowFields><rowItems count="${grid.rowItems.length + 1}">${items(grid.rowItems.length)}</rowItems>`;
	if (col >= 0)
		body += `<colFields count="1"><field x="${col}"/></colFields><colItems count="${grid.columnItems.length + 1}">${items(grid.columnItems.length)}</colItems>`;
	else body += '<colItems count="1"><i><x/></i></colItems>';
	body += `<dataFields count="1"><dataField${attrs({ name: encodeEscapes(pivot.valueCaption), fld: value, subtotal: pivot.aggregate, baseField: 0, baseItem: 0 })}/></dataFields>`;
	const doc = parseXml(
		pivot.sourceXml ??
			`<pivotTableDefinition xmlns="${NS.x}" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" createdVersion="6"><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`,
	);
	const root = doc.documentElement;
	root.setAttribute('name', pivot.name);
	root.setAttribute('cacheId', String(cacheId));
	root.setAttribute('compact', '0');
	root.setAttribute('compactData', '0');
	root.setAttribute('rowGrandTotals', '1');
	root.setAttribute('colGrandTotals', '1');
	const replaced = new Set([
		'location',
		'pivotFields',
		'rowFields',
		'rowItems',
		'colFields',
		'colItems',
		'dataFields',
	]);
	for (const child of elements(root)) if (replaced.has(child.localName!)) root.removeChild(child);
	const replacement = parseXml(`<root xmlns="${NS.x}">${body}</root>`).documentElement;
	const before = root.firstChild;
	for (const child of elements(replacement)) root.insertBefore(doc.importNode(child, true), before);
	return XML_HEADER + buildXml(root);
}
