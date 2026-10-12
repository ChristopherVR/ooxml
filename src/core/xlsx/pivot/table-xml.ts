import { NS, buildXml, elements, first, parseXml } from '../../xml/index';
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
	const second = pivot.secondRowField ? pivot.fields.indexOf(pivot.secondRowField) : -1;
	const col = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const value = pivot.fields.indexOf(pivot.valueField);
	const fields = pivot.fields
		.map((_field, index) => {
			const items =
				index === row
					? grid.rowItems
					: index === second
						? grid.secondRowItems
						: index === col
							? grid.columnItems
							: undefined;
			const subtotal = !(index === row && pivot.subtotals === false);
			const repeated =
				pivot.repeatLabels && (index === row || index === second)
					? '<extLst><ext uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}"><x14:pivotField xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" fillDownLabels="1"/></ext></extLst>'
					: '';
			return `<pivotField${attrs({ axis: index === row || index === second ? 'axisRow' : index === col ? 'axisCol' : undefined, dataField: index === value ? true : undefined, compact: false, outline: false, showAll: false, subtotalTop: index === row && second >= 0 ? false : undefined, defaultSubtotal: subtotal ? undefined : false })}>${items ? `<items count="${items.length + Number(subtotal)}">${items.map((item) => `<item x="${cache.indices[index]!.get(pivotValueKey(item))}"/>`).join('')}${subtotal ? '<item t="default"/>' : ''}</items>` : ''}${repeated}</pivotField>`;
		})
		.join('');
	const items = (count: number) =>
		Array.from({ length: count }, (_v, i) => `<i><x v="${i}"/></i>`).join('') +
		'<i t="grand"><x/></i>';
	const header = grid.firstDataRow - 1;
	const rowIndices = [row, ...(second >= 0 ? [second] : [])];
	let previous: number[] = [];
	const rowItems =
		grid.rowEntries
			?.map((entry) => {
				let prefix = 0;
				if (!entry.type)
					while (prefix < entry.indices.length - 1 && entry.indices[prefix] === previous[prefix])
						prefix++;
				previous = entry.type ? [] : entry.indices;
				return `<i${attrs({ t: entry.type, r: prefix || undefined })}>${entry.indices
					.slice(prefix)
					.map((index) => `<x${attrs({ v: index || undefined })}/>`)
					.join('')}</i>`;
			})
			.join('') ?? items(grid.rowItems.length);
	let body = `<location${attrs({ ref: formatRange(pivot.location), firstHeaderRow: header, firstDataRow: grid.firstDataRow, firstDataCol: rowIndices.length })}/><pivotFields count="${pivot.fields.length}">${fields}</pivotFields><rowFields count="${rowIndices.length}">${rowIndices.map((index) => `<field x="${index}"/>`).join('')}</rowFields><rowItems count="${grid.rowEntries?.length ?? grid.rowItems.length + 1}">${rowItems}</rowItems>`;
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
	for (const extension of elements(
		first(root, 'extLst', NS.x) ?? parseXml('<root/>').documentElement,
	)) {
		const definition = first(
			extension,
			'pivotTableDefinition',
			'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
		);
		if (definition?.hasAttribute('fillDownLabelsDefault'))
			definition.setAttribute('fillDownLabelsDefault', pivot.repeatLabels ? '1' : '0');
	}
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
