import { NS, buildXml, elements, first, parseXml } from '../../xml/index';
import { formatRange } from '../address';
import type { PivotTable } from '../model';
import { XML_HEADER, attrs, encodeEscapes } from '../write/xml-out';
import { pivotGrid } from './aggregate';
import { pivotValueKey } from './data';
import type { PivotCacheData } from './cache-xml';
import { pivotMeasures } from './measures';
import { pivotValueAxisItems } from './values-axis';

/** Writes the supported table structure while retaining unrelated native attributes/extensions. */
export function pivotTableXml(pivot: PivotTable, cacheId: number, cache: PivotCacheData): string {
	const grid = pivotGrid(pivot);
	const row = pivot.fields.indexOf(pivot.rowField);
	const second = pivot.secondRowField ? pivot.fields.indexOf(pivot.secondRowField) : -1;
	const col = pivot.columnField ? pivot.fields.indexOf(pivot.columnField) : -1;
	const measures = pivotMeasures(pivot);
	const multiple = measures.length > 1;
	const rowsAxis = pivot.valuesAxis === 'rows';
	const compact = pivot.layout === 'compact';
	const outline = pivot.layout === 'compact' || pivot.layout === 'outline';
	const top = (pivot.subtotalPosition ?? (outline ? 'top' : 'bottom')) === 'top';
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
			return `<pivotField${attrs({ axis: index === row || index === second ? 'axisRow' : index === col ? 'axisCol' : undefined, dataField: measures.some((measure) => pivot.fields.indexOf(measure.field) === index) ? true : undefined, compact, outline, showAll: false, subtotalTop: index === row ? top : undefined, defaultSubtotal: subtotal ? undefined : false })}>${items ? `<items count="${items.length + Number(subtotal)}">${items.map((item) => `<item x="${cache.indices[index]!.get(pivotValueKey(item))}"/>`).join('')}${subtotal ? '<item t="default"/>' : ''}</items>` : ''}${repeated}</pivotField>`;
		})
		.join('');
	const items = (count: number) =>
		Array.from({ length: count }, (_v, i) => `<i><x v="${i}"/></i>`).join('') +
		'<i t="grand"><x/></i>';
	const header = grid.firstHeaderRow ?? (multiple && col >= 0 ? 1 : grid.firstDataRow - 1);
	const rowIndices =
		rowsAxis && pivot.valuesPosition === 'first'
			? [-2, row, second]
			: rowsAxis && pivot.valuesPosition === 'between'
				? [row, -2, second]
				: [row, ...(second >= 0 ? [second] : []), ...(rowsAxis ? [-2] : [])];
	let previous: number[] = [];
	const rowItems =
		grid.rowEntries
			?.map((entry) => {
				let prefix = 0;
				if (entry.type === 'default' && rowsAxis && pivot.valuesPosition === 'first') prefix = 1;
				else if (!entry.type)
					while (prefix < entry.indices.length - 1 && entry.indices[prefix] === previous[prefix])
						prefix++;
				previous =
					entry.type && !(entry.type === 'default' && pivot.valuesPosition === 'first')
						? []
						: entry.indices;
				return `<i${attrs({ t: entry.type, r: prefix || undefined, i: entry.measure || undefined })}>${entry.indices
					.slice(prefix)
					.map((index) => `<x${attrs({ v: index || undefined })}/>`)
					.join('')}</i>`;
			})
			.join('') ?? items(grid.rowItems.length);
	let body = `<location${attrs({ ref: formatRange(pivot.location), firstHeaderRow: header, firstDataRow: grid.firstDataRow, firstDataCol: grid.rowLabelColumns ?? 1 })}/><pivotFields count="${pivot.fields.length}">${fields}</pivotFields><rowFields count="${rowIndices.length}">${rowIndices.map((index) => `<field x="${index}"/>`).join('')}</rowFields><rowItems count="${grid.rowEntries?.length ?? grid.rowItems.length + 1}">${rowItems}</rowItems>`;
	if (multiple && !rowsAxis)
		body += `<colFields count="${col >= 0 ? 2 : 1}">${col >= 0 ? `<field x="${col}"/>` : ''}<field x="-2"/></colFields>${pivotValueAxisItems(measures.length, col >= 0 ? grid.columnItems.length : undefined)}`;
	else if (col >= 0)
		body += `<colFields count="1"><field x="${col}"/></colFields><colItems count="${grid.columnItems.length + 1}">${items(grid.columnItems.length)}</colItems>`;
	else
		body += rowsAxis
			? '<colItems count="1"><i/></colItems>'
			: '<colItems count="1"><i><x/></i></colItems>';
	body += `<dataFields count="${measures.length}">${measures.map((measure) => `<dataField${attrs({ name: encodeEscapes(measure.caption), fld: pivot.fields.indexOf(measure.field), subtotal: measure.aggregate, baseField: 0, baseItem: 0 })}/>`).join('')}</dataFields>`;
	const doc = parseXml(
		pivot.sourceXml ??
			`<pivotTableDefinition xmlns="${NS.x}" dataCaption="Values" updatedVersion="6" minRefreshableVersion="3" createdVersion="6"><pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`,
	);
	const root = doc.documentElement;
	root.setAttribute('name', pivot.name);
	root.setAttribute('cacheId', String(cacheId));
	root.setAttribute('dataOnRows', rowsAxis ? '1' : '0');
	if (rowsAxis && (pivot.valuesPosition === 'between' || pivot.valuesPosition === 'first'))
		root.setAttribute('dataPosition', pivot.valuesPosition === 'first' ? '0' : '1');
	else root.removeAttribute('dataPosition');
	root.setAttribute('compact', compact ? '1' : '0');
	root.setAttribute('compactData', compact ? '1' : '0');
	root.setAttribute('outline', outline ? '1' : '0');
	root.setAttribute('outlineData', outline ? '1' : '0');
	if (compact) root.setAttribute('indent', '0');
	if (
		multiple &&
		!elements(first(root, 'extLst', NS.x) ?? parseXml('<root/>').documentElement).some(
			(extension) =>
				first(
					extension,
					'pivotTableDefinition',
					'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
				),
		)
	) {
		let list = first(root, 'extLst', NS.x);
		if (!list) {
			list = doc.createElementNS(NS.x, 'extLst');
			root.appendChild(list);
		}
		const extension = parseXml(
			`<ext xmlns="${NS.x}" uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}"><x14:pivotTableDefinition xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" hideValuesRow="1"/></ext>`,
		).documentElement;
		list.appendChild(doc.importNode(extension, true));
	}
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
		if (multiple && definition) definition.setAttribute('hideValuesRow', '1');
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
