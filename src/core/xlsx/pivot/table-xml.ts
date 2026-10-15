import { NS, buildXml, elements, first, parseXml } from '../../xml/index';
import { pivotColumnFields, pivotItems, pivotRowFields, pivotValueKey } from './data';
import { pivotColumnItemsXml, sharedPrefix } from './columns';
import { expandDateGroups } from './date-groups';
import { groupLabelsOf } from './groups';
import { filteredFields, hiddenItemKeys } from './item-filters';
import { formatRange } from '../address';
import type { CellValue, PivotTable } from '../model';
import { XML_HEADER, attrs, encodeEscapes } from '../write/xml-out';
import { pivotGrid } from './aggregate';
import type { PivotCacheData } from './cache-xml';
import { pivotMeasures } from './measures';
import { fieldSubtotalFunctions, PIVOT_SUBTOTALS } from './subtotals';
import { pivotValueAxisItems, remapPivotRowEntries } from './values-axis';

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

/** Writes the supported table structure while retaining unrelated native attributes/extensions. */
export function pivotTableXml(source: PivotTable, cacheId: number, cache: PivotCacheData): string {
	const grid = pivotGrid(source);
	// Date groups add their derived fields after the cache fields; the table lists them all.
	const pivot = expandDateGroups(source);
	const rowNames = pivotRowFields(pivot);
	const rowIdx = rowNames.map((name) => pivot.fields.indexOf(name));
	const row = rowIdx[0]!;
	const second = rowIdx[1] ?? -1;
	const page = pivot.pageFilter ? pivot.fields.indexOf(pivot.pageFilter.field) : -1;
	const pageItems = page >= 0 ? pivotItems(pivot, pivot.pageFilter!.field) : [];
	const filtered = filteredFields(pivot);
	const dateFiltered = pivot.dateFilter !== undefined;
	// Report filters and hidden items keep every cached item in the field's item list.
	const axisItems = (field: string, visible: typeof grid.rowItems): CellValue[] =>
		groupLabelsOf(pivot, field) ??
		(page >= 0 || dateFiltered || filtered.has(field) ? pivotItems(pivot, field) : visible);
	const anyFull =
		page >= 0 ||
		dateFiltered ||
		rowNames.some((name) => filtered.has(name) || groupLabelsOf(pivot, name));
	// The item index of a cached value, or of a group label for a grouped field.
	const itemX = (index: number, item: CellValue) => {
		const labels = groupLabelsOf(pivot, pivot.fields[index]!);
		return labels ? labels.indexOf(String(item)) : cache.indices[index]!.get(pivotValueKey(item));
	};
	const visibleRows = grid.axisItems ?? [
		grid.rowItems,
		...(grid.secondRowItems ? [grid.secondRowItems] : []),
	];
	const rowItemsFull = rowNames.map((name, k) => axisItems(name, visibleRows[k]!));
	const columnNames = pivotColumnFields(pivot);
	const columnIdx = columnNames.map((name) => pivot.fields.indexOf(name));
	const visibleColumns = grid.columnAxisItems ?? (columnNames.length ? [grid.columnItems] : []);
	const columnItemsFull = columnNames.map((name, k) => axisItems(name, visibleColumns[k]!));
	const itemIndices = (visible: typeof grid.rowItems, full: typeof grid.rowItems) => {
		const indices = new Map(full.map((item, index) => [pivotValueKey(item), index]));
		return visible.map((v) => indices.get(pivotValueKey(v))!);
	};
	const rowIndicesFull = rowNames.map((_name, k) => itemIndices(visibleRows[k]!, rowItemsFull[k]!));
	const columnIndicesFull = columnNames.map((_name, k) =>
		itemIndices(visibleColumns[k]!, columnItemsFull[k]!),
	);
	const rowEntries =
		anyFull && grid.rowEntries
			? remapPivotRowEntries(grid.rowEntries, rowIndicesFull)
			: grid.rowEntries;
	const pageOffset = page >= 0 ? 2 : 0;
	const col = columnIdx[0] ?? -1;
	const measures = pivotMeasures(pivot);
	const multiple = measures.length > 1;
	const rowsAxis = pivot.valuesAxis === 'rows';
	const compact = pivot.layout === 'compact';
	const outline = pivot.layout === 'compact' || pivot.layout === 'outline';
	const top = (pivot.subtotalPosition ?? (outline ? 'top' : 'bottom')) === 'top';
	const fields = pivot.fields
		.map((name, index) => {
			const isDate = pivot.dateFields?.includes(name) ?? false;
			const at = rowIdx.indexOf(index);
			const colAt = columnIdx.indexOf(index);
			const items =
				index === page
					? pageItems
					: at >= 0
						? rowItemsFull[at]
						: colAt >= 0
							? columnItemsFull[colAt]
							: (groupLabelsOf(pivot, name) ??
								(filtered.has(name) || isDate ? pivotItems(pivot, name) : undefined));
			const hidden = hiddenItemKeys(pivot, name);
			const functions = at >= 0 ? fieldSubtotalFunctions(pivot, name) : 'auto';
			const automatic = functions === 'auto';
			const custom = automatic ? [] : functions;
			const repeated =
				pivot.repeatLabels && at >= 0
					? `<extLst><ext uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}"><x14:pivotField xmlns:x14="${X14}" fillDownLabels="1"/></ext></extLst>`
					: '';
			const flags = Object.fromEntries(
				custom.map((fn) => [PIVOT_SUBTOTALS[fn].attr, true] as const),
			);
			const entries = automatic
				? '<item t="default"/>'
				: custom.map((fn) => `<item t="${PIVOT_SUBTOTALS[fn].item}"/>`).join('');
			return `<pivotField${attrs({ axis: index === page ? 'axisPage' : at >= 0 ? 'axisRow' : colAt >= 0 ? 'axisCol' : undefined, dataField: measures.some((measure) => pivot.fields.indexOf(measure.field) === index) ? true : undefined, compact, outline, showAll: false, subtotalTop: at >= 0 ? top : undefined, numFmtId: isDate ? 14 : undefined, defaultSubtotal: automatic ? undefined : false, ...flags })}>${items ? `<items count="${items.length + (automatic ? 1 : custom.length)}">${items.map((item) => `<item${hidden.has(pivotValueKey(item)) ? ' h="1"' : ''} x="${itemX(index, item)}"/>`).join('')}${entries}</items>` : ''}${repeated}</pivotField>`;
		})
		.join('');
	const items = (count: number, indices?: readonly number[]) =>
		Array.from({ length: count }, (_v, i) => `<i><x v="${indices?.[i] ?? i}"/></i>`).join('') +
		'<i t="grand"><x/></i>';
	const header =
		grid.firstHeaderRow !== undefined
			? grid.firstHeaderRow - pageOffset
			: col >= 0
				? 1
				: grid.firstDataRow - pageOffset - 1;
	const rowIndices =
		rowsAxis && pivot.valuesPosition === 'first'
			? [-2, row, second]
			: rowsAxis && pivot.valuesPosition === 'between'
				? [row, -2, second]
				: [...rowIdx, ...(rowsAxis ? [-2] : [])];
	let previous: number[] = [];
	const rowItems =
		rowEntries
			?.map((entry) => {
				let prefix = 0;
				const type =
					entry.type === undefined || entry.type === 'default' || entry.type === 'grand'
						? entry.type
						: PIVOT_SUBTOTALS[entry.type].item;
				if (rowsAxis) {
					if (entry.type === 'default' && pivot.valuesPosition === 'first') prefix = 1;
					else if (!entry.type)
						while (prefix < entry.indices.length - 1 && entry.indices[prefix] === previous[prefix])
							prefix++;
					previous =
						entry.type && !(entry.type === 'default' && pivot.valuesPosition === 'first')
							? []
							: entry.indices;
				} else {
					// Values on columns: each item repeats the leading indices it shares with the
					// previous one, whatever its type (a subtotal repeats its group's path).
					if (entry.type !== 'grand') prefix = sharedPrefix(entry.indices, previous);
					previous = entry.type === 'grand' ? [] : entry.indices;
				}
				return `<i${attrs({ t: type, r: prefix || undefined, i: entry.measure || undefined })}>${entry.indices
					.slice(prefix)
					.map((index) => `<x${attrs({ v: index || undefined })}/>`)
					.join('')}</i>`;
			})
			.join('') ?? items(grid.rowItems.length, rowIndicesFull[0]);
	let body = `<location${attrs({ ref: formatRange({ ...pivot.location, start: { ...pivot.location.start, row: pivot.location.start.row + pageOffset } }), firstHeaderRow: header, firstDataRow: grid.firstDataRow - pageOffset, rowPageCount: page >= 0 ? 1 : undefined, colPageCount: page >= 0 ? 1 : undefined, firstDataCol: grid.rowLabelColumns ?? 1 })}/><pivotFields count="${pivot.fields.length}">${fields}</pivotFields><rowFields count="${rowIndices.length}">${rowIndices.map((index) => `<field x="${index}"/>`).join('')}</rowFields><rowItems count="${grid.rowEntries?.length ?? grid.rowItems.length + 1}">${rowItems}</rowItems>`;
	if (!rowsAxis && col >= 0 && grid.columnEntries) {
		const fieldsOnColumns = [...columnIdx, ...(multiple ? [-2] : [])];
		body += `<colFields count="${fieldsOnColumns.length}">${fieldsOnColumns.map((index) => `<field x="${index}"/>`).join('')}</colFields>${pivotColumnItemsXml(grid.columnEntries, measures.length, (level, index) => columnIndicesFull[level]![index]!)}`;
	} else if (multiple && !rowsAxis)
		body += `<colFields count="1"><field x="-2"/></colFields>${pivotValueAxisItems(measures.length)}`;
	else if (col >= 0)
		body += `<colFields count="1"><field x="${col}"/></colFields><colItems count="${grid.columnItems.length + 1}">${items(grid.columnItems.length, columnIndicesFull[0])}</colItems>`;
	else
		body += rowsAxis
			? '<colItems count="1"><i/></colItems>'
			: page >= 0
				? '<colItems count="1"><i/></colItems>'
				: '<colItems count="1"><i><x/></i></colItems>';
	if (page >= 0)
		body += `<pageFields count="1"><pageField${attrs({ fld: page, hier: -1, item: pivot.pageFilter!.item !== undefined ? pageItems.findIndex((item) => pivotValueKey(item) === pivotValueKey(pivot.pageFilter!.item!)) : undefined })}/></pageFields>`;
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
		'pageFields',
		'filters',
	]);
	for (const child of elements(root)) if (replaced.has(child.localName!)) root.removeChild(child);
	const replacement = parseXml(`<root xmlns="${NS.x}">${body}</root>`).documentElement;
	const before = root.firstChild;
	for (const child of elements(replacement)) root.insertBefore(doc.importNode(child, true), before);
	if (pivot.dateFilter) {
		// `filters` follows the table style and precedes the extension list.
		const filters = parseXml(
			`<root xmlns="${NS.x}">${dateFilterXml(pivot, pivot.dateFilter)}</root>`,
		).documentElement;
		root.insertBefore(
			doc.importNode(elements(filters)[0]!, true),
			first(root, 'extLst', NS.x) ?? null,
		);
	}
	return XML_HEADER + buildXml(root);
}

/** The timeline's `filters` element: one inclusive date range on a date field. */
function dateFilterXml(pivot: PivotTable, filter: NonNullable<PivotTable['dateFilter']>): string {
	const x15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
	return (
		`<filters count="1"><filter${attrs({ fld: pivot.fields.indexOf(filter.field), type: 'dateBetween', evalOrder: -1, id: 1, name: filter.field })}>` +
		`<autoFilter ref="A1"><filterColumn colId="0"><customFilters and="1">` +
		`<customFilter operator="greaterThanOrEqual" val="${Math.floor(filter.start)}"/>` +
		`<customFilter operator="lessThanOrEqual" val="${Math.floor(filter.end)}"/>` +
		`</customFilters></filterColumn></autoFilter>` +
		`<extLst><ext uri="{0605FD5F-26C8-4aeb-8148-2DB25E43C511}" xmlns:x15="${x15}"><x15:pivotFilter useWholeDay="1"/></ext></extLst></filter></filters>`
	);
}
