import { NS, children, elements, first, parseXml } from '../../xml/index';
import { parseRange } from '../address';
import type { PivotTable, Workbook } from '../model';
import type { SourceIndex } from '../read/package';
import { att, decodeEscapes } from '../read/xml-util';
import { PIVOT_AGGREGATES } from './data';
import { readPivotCaptions } from './captions-read';
import { pivotCaptions, trimmedCaptions } from './captions';
import { readDateFields, readDateFilter } from './date-read';
import { readFieldGroups } from './group-read';
import { readDateGroups } from './date-group-read';
import { fieldSubtotalsOf } from './subtotal-read';
import { readItemFilters, readSlicerCacheId } from './item-filters-read';
import { readItemNames } from './item-names-read';
import { assertPageAxisItems } from './page-filter-admission';
import { readPivotPageFilter } from './page-filter';
import { supportedPivotAxes, pivotFieldIndex, pivotValueIndex } from './values-axis';
import { readPivotDisplay } from './display';
import { readPivotCondition } from './condition';
import { assertPivotFieldPlacement, readPivotCalculatedField } from './calculated';
import { readPivotCache } from './read-cache';
import { pivotNumberFormat } from './number-format-read';

export function readPivotDefinition(
	workbook: Workbook,
	source: SourceIndex,
	part: string,
): PivotTable {
	const xml = source.text(part)!;
	const root = parseXml(xml).documentElement;
	const {
		cacheId,
		cache,
		reference,
		recordRoot,
		cacheFieldCount,
		formula,
		groupedFields,
		databaseCount,
		fields,
		shared,
		records,
	} = readPivotCache(workbook, source, part, root);
	const rowsAxis = att(root, 'dataOnRows') === '1';
	const rowFields = children(
		first(root, 'rowFields', NS.x) ?? parseXml('<root/>').documentElement,
		'field',
		NS.x,
	);
	const colFields = children(
		first(root, 'colFields', NS.x) ?? parseXml('<root/>').documentElement,
		'field',
		NS.x,
	);
	const dataFields = children(
		first(root, 'dataFields', NS.x) ?? parseXml('<root/>').documentElement,
		'dataField',
		NS.x,
	);
	const calculated = readPivotCalculatedField(formula, dataFields, cacheFieldCount);
	if (
		rowFields.length < 1 ||
		dataFields.length < 1 ||
		dataFields.length > 16 ||
		!supportedPivotAxes(rowFields, colFields, dataFields.length, rowsAxis)
	)
		throw new Error('Unsupported pivot layout');
	const valueIndex = pivotValueIndex(root, rowsAxis ? rowFields : colFields);
	const measure = dataFields[0]!;
	const definition = elements(first(root, 'extLst', NS.x) ?? parseXml('<root/>').documentElement)
		.map((extension) =>
			first(
				extension,
				'pivotTableDefinition',
				'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
			),
		)
		.find(Boolean);
	if (dataFields.length > 1 && att(definition, 'hideValuesRow') !== '1')
		throw new Error('Visible Values header rows are preserved');
	for (const [axis, count] of [
		['rowItems', rowsAxis ? dataFields.length : 1],
		['colItems', rowsAxis ? 1 : dataFields.length],
	] as const) {
		for (const item of children(
			first(root, axis, NS.x) ?? parseXml('<root/>').documentElement,
			'i',
			NS.x,
		)) {
			const index = Number(att(item, 'i') ?? '0');
			if (!Number.isInteger(index) || index < 0 || index >= count)
				throw new Error('Unsupported Values axis item');
		}
	}
	const aggregate = (att(measure, 'subtotal') ?? 'sum') as PivotTable['aggregate'];
	const valueNumberFormat = pivotNumberFormat(source, att(measure, 'numFmtId'));
	if (
		!PIVOT_AGGREGATES.includes(aggregate) ||
		(att(measure, 'showDataAs') && att(measure, 'showDataAs') !== 'normal')
	)
		throw new Error('Unsupported pivot aggregate');
	const names = fields.map((field) => decodeEscapes(att(field, 'name') ?? ''));
	const axisNames = groupedFields.map((field) => decodeEscapes(att(field, 'name') ?? ''));
	const realRows = rowsAxis
		? rowFields.filter((field) => pivotFieldIndex(field) !== -2)
		: rowFields;
	const row = axisNames[pivotFieldIndex(realRows[0])];
	const second = realRows[1] ? axisNames[Number(att(realRows[1], 'x'))] : undefined;
	const more = realRows.slice(2).map((field) => axisNames[Number(att(field, 'x'))]);
	const pageFilter = readPivotPageFilter(root, names, shared);
	const tableFields = children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x);
	assertPivotFieldPlacement(tableFields, calculated ? cacheFieldCount - 1 : -1);
	const outer = tableFields[pivotFieldIndex(realRows[0])];
	const compact = att(outer, 'compact') !== '0';
	const outline = att(outer, 'outline') !== '0';
	const layout = compact ? 'compact' : outline ? 'outline' : 'tabular';
	const defaultCompact =
		compact &&
		outline &&
		att(root, 'outline') === undefined &&
		att(root, 'outlineData') === undefined;
	if (
		tableFields.some(
			(field) =>
				(att(field, 'compact') !== '0') !== compact || (att(field, 'outline') !== '0') !== outline,
		) ||
		(att(root, 'compact') !== '0') !== compact ||
		(att(root, 'compactData') !== '0') !== compact ||
		(!defaultCompact &&
			((att(root, 'outline') === '1') !== outline ||
				(att(root, 'outlineData') === '1') !== outline)) ||
		(layout === 'compact' && att(root, 'indent') !== '0')
	)
		throw new Error('Mixed field layouts or custom compact indentation are preserved');
	const rowDefinitions = realRows.map((field) => tableFields[Number(att(field, 'x'))]!);
	const subtotalSettings = fieldSubtotalsOf(rowDefinitions, [row!, second, ...more]);
	const subtotal = !subtotalSettings.some(
		(entry) => entry.field === row && entry.functions.length === 0,
	);
	const repeat = realRows.map((field) => {
		const node = tableFields[Number(att(field, 'x'))];
		const extension = first(first(node, 'extLst', NS.x), 'ext', NS.x);
		return (
			att(
				first(
					extension,
					'pivotField',
					'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main',
				),
				'fillDownLabels',
			) === '1'
		);
	});
	if (repeat.some((entry) => entry !== repeat[0])) throw new Error('Mixed field label repetition');
	// With several measures Values (-2) closes the column fields.
	const columnDefs = dataFields.length > 1 && !rowsAxis ? colFields.slice(0, -1) : colFields;
	const columnNames = columnDefs.map((field) => axisNames[Number(att(field, 'x'))]);
	const hasColumn = columnDefs.length > 0;
	const col = columnNames[0];
	const moreColumns = columnNames.slice(1);
	const columnDefinitions = columnDefs.map((field) => tableFields[Number(att(field, 'x'))]!);
	if (columnNames.some((name) => !name)) throw new Error('Invalid pivot column fields');
	const columnSubtotals = fieldSubtotalsOf(columnDefinitions, columnNames);
	const data = names[Number(att(measure, 'fld'))];
	const location = parseRange(att(first(root, 'location', NS.x), 'ref') ?? '');
	if (!row || !data || !location || (hasColumn && !col) || (realRows.length === 2 && !second))
		throw new Error('Invalid pivot fields');
	if (pageFilter) {
		if (location.start.row < 2) throw new Error('Invalid report filter location');
		location.start.row -= 2;
	}
	const itemFilters = readItemFilters(tableFields, names, shared);
	const itemNames = readItemNames(tableFields, names, shared);
	const slicerCacheId = readSlicerCacheId(cache);
	const fieldGroups = readFieldGroups(fields, names);
	const dateFields = readDateFields(fields, names);
	tableFields.forEach((field, index) => {
		const format = att(field, 'numFmtId');
		if (format !== undefined && format !== '0' && !dateFields.includes(names[index] ?? ''))
			throw new Error('Unsupported pivot field number format');
	});
	const dateGroups = readDateGroups(
		groupedFields,
		databaseCount,
		names,
		dateFields,
		records,
		workbook.date1904,
	);
	if (
		fieldGroups.some(
			(group) => group.field === pageFilter?.field || dateFields.includes(group.field),
		)
	)
		throw new Error('A grouped field cannot be a report filter or date field');
	const condition = readPivotCondition(root, pivotFieldIndex(realRows[0]));
	const dateFilter = condition
		? undefined
		: readDateFilter(root, names, dateFields, workbook.date1904);
	const display = readPivotDisplay(root, tableFields, rowDefinitions, rowsAxis);
	const onAxis = new Set<string | undefined>([
		row,
		second,
		...more,
		...columnNames,
		pageFilter?.field,
		...dataFields.map((field) => names[Number(att(field, 'fld'))]),
	]);
	if (dateFields.some((field) => onAxis.has(field)))
		throw new Error('Date fields used as rows, columns or values are preserved');
	// The first row field's None is the `subtotals` flag; every other setting is an entry.
	const fieldSubtotals = [
		...subtotalSettings.filter((entry) => entry.field !== row || entry.functions.length),
		...columnSubtotals,
	];
	const pivot: PivotTable = {
		...(calculated ? { calculatedField: calculated } : {}),
		...(display ? { display } : {}),
		...(condition ? { condition } : {}),
		...(pageFilter ? { pageFilter } : {}),
		...(itemFilters ? { itemFilters } : {}),
		...(itemNames.length ? { itemNames } : {}),
		...(fieldGroups.length ? { fieldGroups } : {}),
		...(dateFields.length ? { dateFields } : {}),
		...(dateGroups.length ? { dateGroups } : {}),
		...(dateGroups.length && workbook.date1904 ? { date1904: true } : {}),
		...(dateFilter ? { dateFilter } : {}),
		...(slicerCacheId !== undefined ? { slicerCacheId } : {}),
		name: att(root, 'name') ?? 'PivotTable',
		source: reference,
		rowField: row,
		layout,
		...(rowsAxis ? { valuesAxis: 'rows' as const } : {}),
		...(rowsAxis && valueIndex < realRows.length
			? {
					valuesPosition:
						valueIndex === 0
							? ('first' as const)
							: valueIndex === 1 && realRows.length === 2
								? ('between' as const)
								: valueIndex,
				}
			: {}),
		subtotalPosition: att(outer, 'subtotalTop') === '0' ? 'bottom' : 'top',
		...(fieldSubtotals.length ? { fieldSubtotals } : {}),
		...(more.length ? { moreRowFields: more as string[] } : {}),
		...(second
			? { secondRowField: second, subtotals: subtotal, repeatLabels: repeat[0] ?? false }
			: {
					...(subtotal ? {} : { subtotals: false }),
					...(repeat[0] ? { repeatLabels: true } : {}),
				}),
		valueField: data,
		aggregate,
		valueCaption: decodeEscapes(att(measure, 'name') ?? data),
		...(valueNumberFormat ? { valueNumberFormat } : {}),
		...(dataFields.length > (calculated ? 2 : 1)
			? {
					additionalValues: dataFields.slice(1, calculated ? -1 : undefined).map((measure) => {
						const field = names[Number(att(measure, 'fld'))];
						const aggregate = (att(measure, 'subtotal') ?? 'sum') as PivotTable['aggregate'];
						if (
							!field ||
							!PIVOT_AGGREGATES.includes(aggregate) ||
							(att(measure, 'showDataAs') && att(measure, 'showDataAs') !== 'normal')
						)
							throw new Error('Unsupported pivot measure');
						const numberFormat = pivotNumberFormat(source, att(measure, 'numFmtId'));
						return {
							field,
							aggregate,
							caption: decodeEscapes(att(measure, 'name') ?? field),
							...(numberFormat ? { numberFormat } : {}),
						};
					}),
				}
			: {}),
		fields: names,
		records,
		location,
		partName: part,
		cacheId,
		sourceXml: xml,
		...(col ? { columnField: col } : {}),
		...(moreColumns.length ? { moreColumnFields: moreColumns as string[] } : {}),
	};
	// The header, grand total and Values captions the definition names; the cached cells add the rest.
	const captions = trimmedCaptions(
		pivotCaptions({
			captions: readPivotCaptions(root, pivot),
		}),
	);
	if (captions) pivot.captions = captions;
	if (pageFilter && moreColumns.length) throw new Error('Report filters support one column field');
	assertPageAxisItems(root, pivot, shared, cache, recordRoot);
	return pivot;
}
