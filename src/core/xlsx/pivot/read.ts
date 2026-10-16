import { RELATIONSHIP_TYPES } from '../../opc/index';
import { NS, children, elements, first, parseXml, relAttr, type XmlElement } from '../../xml/index';
import { parseRange, quoteSheetName, formatRange } from '../address';
import { getCell } from '../cells';
import { isErrorCode, type CellValue, type PivotTable, type Workbook } from '../model';
import type { SourceIndex } from '../read/package';
import { att, decodeEscapes } from '../read/xml-util';
import { PIVOT_AGGREGATES } from './data';
import { pivotGrid } from './aggregate';
import { assertPivotAdmission, assertPivotRecords } from './admission';
import {
	aggregateCaptionsOf,
	inferPivotCaptions,
	readPivotCaptionAttributes,
} from './captions-read';
import { pivotCaptions, trimmedCaptions } from './captions';
import { pivotKey, type SlicerRefs } from '../slicer/read';
import { readDateFields, readDateFilter } from './date-read';
import { readFieldGroups } from './group-read';
import { databaseFieldCount, readDateGroups } from './date-group-read';
import { fieldSubtotalsOf } from './subtotal-read';
import { isoToSerial } from './dates';
import { readItemFilters, readSlicerCacheId } from './item-filters-read';
import { readItemNames } from './item-names-read';
import { assertPageAxisItems } from './page-filter-admission';
import { readPivotPageFilter } from './page-filter';
import { supportedPivotAxes, pivotFieldIndex, pivotValueIndex } from './values-axis';

function value(node: XmlElement, date1904 = false): CellValue {
	const v = att(node, 'v') ?? '';
	if (node.localName === 'd') {
		const serial = isoToSerial(v, date1904);
		if (serial === undefined) throw new Error('Unsupported pivot date item');
		return serial;
	}
	if (node.localName === 'm') return null;
	if (node.localName === 's') return decodeEscapes(v);
	if (node.localName === 'b') return v === '1';
	if (node.localName === 'n' && v && Number.isFinite(Number(v))) return Number(v);
	if (node.localName === 'e' && isErrorCode(v)) return { error: v };
	throw new Error('Unsupported pivot cache item');
}

const namedAggregates = (pivot: PivotTable): PivotTable['captions'] => {
	const aggregates = aggregateCaptionsOf(pivot);
	return aggregates ? { aggregates } : {};
};

function readOne(workbook: Workbook, source: SourceIndex, part: string): PivotTable {
	const xml = source.text(part)!;
	const root = parseXml(xml).documentElement;
	const cacheId = Number(att(root, 'cacheId'));
	if (!Number.isInteger(cacheId) || cacheId < 0) throw new Error('Invalid pivot cache identity');
	const cachePart = source.targetOfType(part, RELATIONSHIP_TYPES.pivotCacheDefinition);
	if (!cachePart) throw new Error('Missing pivot cache');
	const bookPart = source.workbookPart()!;
	const entry = children(
		first(parseXml(source.text(bookPart)!).documentElement, 'pivotCaches', NS.x)!,
		'pivotCache',
		NS.x,
	).find((node) => Number(att(node, 'cacheId')) === cacheId);
	const relationship = entry && source.rels(bookPart).get(relAttr(entry, 'id') ?? '');
	if (!relationship || source.target(bookPart, relationship) !== cachePart)
		throw new Error('Inconsistent pivot cache identity');
	const cache = parseXml(source.text(cachePart)!).documentElement;
	assertPivotAdmission(root, cache);
	if (source.rels(part).size !== 1 || source.rels(cachePart).size !== 1)
		throw new Error('Unsupported PivotTable relationships');
	const sourceNode = first(first(cache, 'cacheSource', NS.x), 'worksheetSource', NS.x);
	const dataSheet = sourceNode && att(sourceNode, 'sheet');
	const dataRange = sourceNode && parseRange(att(sourceNode, 'ref') ?? '');
	const recordPart = source.targetOfType(cachePart, `${NS.r}/pivotCacheRecords`);
	const rowsAxis = att(root, 'dataOnRows') === '1';
	if (
		!dataSheet ||
		!dataRange ||
		!recordPart ||
		!workbook.sheets.some((sheet) => sheet.name === dataSheet)
	)
		throw new Error('Unsupported pivot source');
	const recordRoot = parseXml(source.text(recordPart)!).documentElement;
	assertPivotRecords(recordRoot);
	const allFields = children(first(cache, 'cacheFields', NS.x)!, 'cacheField', NS.x);
	if (allFields.some((field) => field.hasAttribute('formula')))
		throw new Error('Calculated cache field');
	// Date groups append derived cache fields that have no records.
	const databaseCount = databaseFieldCount(allFields);
	const fields = allFields.slice(0, databaseCount);
	const shared = fields.map((field) =>
		elements(first(field, 'sharedItems', NS.x) ?? parseXml('<root/>').documentElement).map((node) =>
			value(node, workbook.date1904),
		),
	);
	const records = children(recordRoot, 'r', NS.x).map((record) =>
		elements(record).map((node, field): CellValue => {
			if (node.localName !== 'x') return value(node, workbook.date1904);
			const index = Number(att(node, 'v'));
			if (!Number.isInteger(index) || index < 0 || index >= (shared[field]?.length ?? 0))
				throw new Error('Invalid pivot cache index');
			return shared[field]![index]!;
		}),
	);
	if (
		records.some((record) => record.length !== fields.length) ||
		records.length * fields.length > 1_000_000
	)
		throw new Error('Invalid pivot records');
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
	if (
		rowFields.length < 1 ||
		dataFields.length < 1 ||
		dataFields.length > 16 ||
		!supportedPivotAxes(rowFields, colFields, dataFields.length, rowsAxis) ||
		root.getAttribute('rowGrandTotals') === '0' ||
		root.getAttribute('colGrandTotals') === '0'
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
	if (
		!PIVOT_AGGREGATES.includes(aggregate) ||
		(att(measure, 'showDataAs') && att(measure, 'showDataAs') !== 'normal')
	)
		throw new Error('Unsupported pivot aggregate');
	const names = fields.map((field) => decodeEscapes(att(field, 'name') ?? ''));
	const axisNames = allFields.map((field) => decodeEscapes(att(field, 'name') ?? ''));
	const realRows = rowsAxis
		? rowFields.filter((field) => pivotFieldIndex(field) !== -2)
		: rowFields;
	const row = axisNames[pivotFieldIndex(realRows[0])];
	const second = realRows[1] ? axisNames[Number(att(realRows[1], 'x'))] : undefined;
	const more = realRows.slice(2).map((field) => axisNames[Number(att(field, 'x'))]);
	const pageFilter = readPivotPageFilter(root, names, shared);
	const tableFields = children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x);
	const outer = tableFields[pivotFieldIndex(realRows[0])];
	const compact = att(outer, 'compact') !== '0';
	const outline = att(outer, 'outline') !== '0';
	const layout = compact ? 'compact' : outline ? 'outline' : 'tabular';
	if (
		tableFields.some(
			(field) =>
				(att(field, 'compact') !== '0') !== compact || (att(field, 'outline') !== '0') !== outline,
		) ||
		(att(root, 'compact') !== '0') !== compact ||
		(att(root, 'compactData') !== '0') !== compact ||
		(att(root, 'outline') === '1') !== outline ||
		(att(root, 'outlineData') === '1') !== outline ||
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
		allFields,
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
	const dateFilter = readDateFilter(root, names, dateFields, workbook.date1904);
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
		source: `${quoteSheetName(dataSheet)}!${formatRange(dataRange)}`,
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
		...(dataFields.length > 1
			? {
					additionalValues: dataFields.slice(1).map((measure) => {
						const field = names[Number(att(measure, 'fld'))];
						const aggregate = (att(measure, 'subtotal') ?? 'sum') as PivotTable['aggregate'];
						if (
							!field ||
							!PIVOT_AGGREGATES.includes(aggregate) ||
							(att(measure, 'showDataAs') && att(measure, 'showDataAs') !== 'normal')
						)
							throw new Error('Unsupported pivot measure');
						return { field, aggregate, caption: decodeEscapes(att(measure, 'name') ?? field) };
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
			captions: { ...readPivotCaptionAttributes(root), ...namedAggregates(pivot) },
		}),
	);
	if (captions) pivot.captions = captions;
	if (pageFilter && moreColumns.length) throw new Error('Report filters support one column field');
	assertPageAxisItems(root, pivot, shared, cache, recordRoot);
	return pivot;
}

/** Admits the supported layout only when its cached cell output agrees with its cache. */
export function readPivotTables(
	workbook: Workbook,
	source: SourceIndex,
	warn: (message: string) => void,
	slicers: SlicerRefs,
): void {
	const caches = new Map<string, number>();
	for (const sheet of workbook.sheets)
		if (sheet.partName)
			for (const rel of source.rels(sheet.partName).values()) {
				if (rel.type !== RELATIONSHIP_TYPES.pivotTable) continue;
				const part = source.target(sheet.partName, rel);
				const cache = part && source.targetOfType(part, RELATIONSHIP_TYPES.pivotCacheDefinition);
				if (cache) caches.set(cache, (caches.get(cache) ?? 0) + 1);
			}
	for (const sheet of workbook.sheets)
		if (sheet.partName) {
			const pivots: PivotTable[] = [];
			for (const rel of source.rels(sheet.partName).values()) {
				if (rel.type !== RELATIONSHIP_TYPES.pivotTable) continue;
				const part = source.target(sheet.partName, rel);
				if (!part) continue;
				try {
					const cache = source.targetOfType(part, RELATIONSHIP_TYPES.pivotCacheDefinition);
					const name = att(parseXml(source.text(part) ?? '<x/>').documentElement, 'name') ?? '';
					if (
						slicers.blocked.has(pivotKey(sheet.sheetId, name)) ||
						(cache && (caches.get(cache) ?? 0) > 1)
					)
						throw new Error('Shared pivot cache or unsupported slicer connection');
					const pivot = readOne(workbook, source, part);
					// A localized Excel wrote its words into the cells; reproduce them on refresh.
					const inferred = inferPivotCaptions(pivot, sheet, {
						...readPivotCaptionAttributes(parseXml(pivot.sourceXml!).documentElement),
						...namedAggregates(pivot),
					});
					if (inferred) pivot.captions = inferred;
					else delete pivot.captions;
					const grid = pivotGrid(pivot);
					if (
						grid.values.length !== pivot.location.end.row - pivot.location.start.row + 1 ||
						grid.values[0]!.length !== pivot.location.end.col - pivot.location.start.col + 1 ||
						grid.values.some((line, r) =>
							line.some(
								(value, c) =>
									JSON.stringify(
										getCell(sheet, pivot.location.start.row + r, pivot.location.start.col + c)
											?.value ?? null,
									) !== JSON.stringify(value),
							),
						)
					)
						throw new Error('Unsupported cached layout');
					pivots.push(pivot);
				} catch {
					const xml = source.text(part);
					if (xml) {
						const name = att(parseXml(xml).documentElement, 'name');
						if (name)
							sheet.preserved.set('source:unsupportedPivotNames', [
								...(sheet.preserved.get('source:unsupportedPivotNames') ?? []),
								name,
							]);
					}
					warn(
						`PivotTable parts on "${sheet.name}" are preserved; this layout or data source cannot be refreshed or configured.`,
					);
				}
			}
			if (pivots.length) {
				sheet.pivotTables = pivots;
				sheet.preserved.set(
					'source:pivotTables',
					pivots.map((pivot) => pivot.partName!),
				);
			}
		}
}
