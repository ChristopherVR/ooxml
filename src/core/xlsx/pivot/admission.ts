import { NS, elements, type XmlElement } from '../../xml/index';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const LAYOUT = 'http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout';
const TABLE: Record<string, string> = {
	pivotTableDefinition:
		'name cacheId dataCaption updatedVersion minRefreshableVersion createdVersion applyNumberFormats applyBorderFormats applyFontFormats applyPatternFormats applyAlignmentFormats applyWidthHeightFormats useAutoFormatting itemPrintTitles indent compact compactData outline outlineData multipleFieldFilters rowGrandTotals colGrandTotals',
	location: 'ref firstHeaderRow firstDataRow firstDataCol',
	pivotFields: 'count',
	pivotField: 'axis dataField compact outline showAll subtotalTop defaultSubtotal',
	items: 'count',
	item: 'x t',
	rowFields: 'count',
	colFields: 'count',
	field: 'x',
	rowItems: 'count',
	colItems: 'count',
	i: 't r',
	x: 'v',
	dataFields: 'count',
	dataField: 'name fld subtotal baseField baseItem showDataAs',
	pivotTableStyleInfo:
		'name showRowHeaders showColHeaders showRowStripes showColStripes showLastColumn',
	extLst: '',
	ext: 'uri',
};
const CACHE: Record<string, string> = {
	pivotCacheDefinition:
		'id refreshedBy refreshedDate createdVersion refreshedVersion minRefreshableVersion recordCount',
	cacheSource: 'type',
	worksheetSource: 'ref sheet',
	cacheFields: 'count',
	cacheField: 'name numFmtId',
	sharedItems:
		'count containsString containsNumber containsInteger containsBlank containsMixedTypes containsSemiMixedTypes minValue maxValue',
	s: 'v',
	n: 'v',
	b: 'v',
	e: 'v',
	m: '',
	extLst: '',
	ext: 'uri',
};

function inspect(node: XmlElement, profile: Record<string, string>, cache: boolean): void {
	if (node.namespaceURI !== NS.x) {
		const owner = node.parentNode?.parentNode?.parentNode as XmlElement | null | undefined;
		const known = cache
			? node.namespaceURI === X14 &&
				node.localName === 'pivotCacheDefinition' &&
				owner?.localName === 'pivotCacheDefinition'
			: (node.namespaceURI === X14 &&
					node.localName === 'pivotTableDefinition' &&
					owner?.localName === 'pivotTableDefinition') ||
				(node.namespaceURI === X14 &&
					node.localName === 'pivotField' &&
					owner?.localName === 'pivotField') ||
				(node.namespaceURI === LAYOUT &&
					node.localName === 'pivotTableDefinition16' &&
					owner?.localName === 'pivotTableDefinition');
		if (!known || elements(node).length) throw new Error('Unsupported PivotTable extension');
		for (let i = 0; i < node.attributes.length; i++) {
			const attr = node.attributes.item(i)!;
			if (attr.name.startsWith('xmlns')) continue;
			if (
				!cache &&
				node.namespaceURI === X14 &&
				node.localName === 'pivotTableDefinition' &&
				attr.name === 'hideValuesRow' &&
				attr.value === '1'
			)
				continue;
			if (
				!cache &&
				node.namespaceURI === X14 &&
				((node.localName === 'pivotField' && attr.name === 'fillDownLabels') ||
					(node.localName === 'pivotTableDefinition' && attr.name === 'fillDownLabelsDefault')) &&
				['0', '1'].includes(attr.value)
			)
				continue;
			throw new Error('Unsupported PivotTable extension option');
		}
		return;
	}
	const allowed = profile[node.localName!]?.split(' ');
	if (!allowed) throw new Error('Unsupported PivotTable markup');
	if (node.localName === 'ext') {
		const uris = cache
			? ['{725AE2AE-9491-48be-B2B4-4EB974FC3084}']
			: [
					'{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}',
					'{747A6164-185A-40DC-8AA5-F01512510D54}',
					'{2946ED86-A175-432a-8AC1-64E0C546D7DE}',
				];
		if (!uris.includes(node.getAttribute('uri') ?? ''))
			throw new Error('Unsupported PivotTable extension identity');
	}
	if (node.localName === 'cacheSource' && node.getAttribute('type') !== 'worksheet')
		throw new Error('Unsupported PivotTable data source');
	for (let i = 0; i < node.attributes.length; i++) {
		const attr = node.attributes.item(i)!;
		if (
			attr.name.startsWith('xmlns') ||
			(attr.namespaceURI === NS.mc && attr.localName === 'Ignorable') ||
			(attr.namespaceURI === 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision' &&
				attr.localName === 'uid')
		)
			continue;
		if (
			!allowed.includes(attr.localName!) ||
			(attr.namespaceURI &&
				!(
					cache &&
					node.localName === 'pivotCacheDefinition' &&
					attr.namespaceURI === NS.r &&
					attr.localName === 'id'
				))
		)
			throw new Error('Unsupported PivotTable option');
		if (attr.localName === 'numFmtId' && attr.value !== '0')
			throw new Error('Unsupported PivotTable field number format');
		if (
			node.localName === 'pivotTableDefinition' &&
			['compact', 'compactData', 'outline', 'outlineData'].includes(attr.localName!) &&
			!['0', '1'].includes(attr.value)
		)
			throw new Error('Unsupported PivotTable global layout');
		if (
			node.localName === 'item' &&
			attr.localName === 't' &&
			!['default', 'data'].includes(attr.value)
		)
			throw new Error('Unsupported PivotTable subtotal');
		if (
			node.localName === 'pivotField' &&
			['subtotalTop', 'defaultSubtotal'].includes(attr.localName!) &&
			!['0', '1'].includes(attr.value)
		)
			throw new Error('Unsupported PivotTable subtotal option');
		if (
			node.localName === 'pivotField' &&
			((attr.localName === 'showAll' && attr.value !== '0') ||
				(['compact', 'outline'].includes(attr.localName!) && !['0', '1'].includes(attr.value)))
		)
			throw new Error('Unsupported PivotTable field layout');
	}
	for (const child of elements(node)) inspect(child, profile, cache);
}

/** Only the explicitly modeled profile and inert native layout extensions are editable. */
export function assertPivotAdmission(table: XmlElement, cache: XmlElement): void {
	inspect(table, TABLE, false);
	inspect(cache, CACHE, true);
}

export function assertPivotRecords(records: XmlElement): void {
	inspect(
		records,
		{ pivotCacheRecords: 'count', r: '', x: 'v', s: 'v', n: 'v', b: 'v', e: 'v', m: '' },
		true,
	);
}
