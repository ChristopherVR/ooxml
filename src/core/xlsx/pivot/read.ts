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

function value(node: XmlElement): CellValue {
	const v = att(node, 'v') ?? '';
	if (node.localName === 'm') return null;
	if (node.localName === 's') return decodeEscapes(v);
	if (node.localName === 'b') return v === '1';
	if (node.localName === 'n' && v && Number.isFinite(Number(v))) return Number(v);
	if (node.localName === 'e' && isErrorCode(v)) return { error: v };
	throw new Error('Unsupported pivot cache item');
}

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
	if (
		!dataSheet ||
		!dataRange ||
		!recordPart ||
		!workbook.sheets.some((sheet) => sheet.name === dataSheet)
	)
		throw new Error('Unsupported pivot source');
	assertPivotRecords(parseXml(source.text(recordPart)!).documentElement);
	const fields = children(first(cache, 'cacheFields', NS.x)!, 'cacheField', NS.x);
	if (fields.some((field) => first(field, 'fieldGroup', NS.x) || field.hasAttribute('formula')))
		throw new Error('Grouped or calculated cache field');
	const shared = fields.map((field) =>
		elements(first(field, 'sharedItems', NS.x) ?? parseXml('<root/>').documentElement).map(value),
	);
	const records = children(parseXml(source.text(recordPart)!).documentElement, 'r', NS.x).map(
		(record) =>
			elements(record).map((node, field): CellValue => {
				if (node.localName !== 'x') return value(node);
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
		rowFields.length > 2 ||
		colFields.length > 1 ||
		dataFields.length !== 1 ||
		first(root, 'pageFields', NS.x) ||
		first(root, 'filters', NS.x) ||
		root.getAttribute('rowGrandTotals') === '0' ||
		root.getAttribute('colGrandTotals') === '0'
	)
		throw new Error('Unsupported pivot layout');
	const measure = dataFields[0]!;
	const aggregate = (att(measure, 'subtotal') ?? 'sum') as PivotTable['aggregate'];
	if (
		!PIVOT_AGGREGATES.includes(aggregate) ||
		(att(measure, 'showDataAs') && att(measure, 'showDataAs') !== 'normal')
	)
		throw new Error('Unsupported pivot aggregate');
	const names = fields.map((field) => decodeEscapes(att(field, 'name') ?? ''));
	const row = names[Number(att(rowFields[0], 'x'))];
	const second = rowFields[1] && names[Number(att(rowFields[1], 'x'))];
	const tableFields = children(first(root, 'pivotFields', NS.x)!, 'pivotField', NS.x);
	const outer = tableFields[Number(att(rowFields[0], 'x'))];
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
	const subtotal = att(outer, 'defaultSubtotal') !== '0';
	const repeat = rowFields.map((field) => {
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
	const col = colFields[0] && names[Number(att(colFields[0], 'x'))];
	const data = names[Number(att(measure, 'fld'))];
	const location = parseRange(att(first(root, 'location', NS.x), 'ref') ?? '');
	if (
		!row ||
		!data ||
		!location ||
		(colFields.length && !col) ||
		(rowFields.length === 2 && !second)
	)
		throw new Error('Invalid pivot fields');
	const pivot: PivotTable = {
		name: att(root, 'name') ?? 'PivotTable',
		source: `${quoteSheetName(dataSheet)}!${formatRange(dataRange)}`,
		rowField: row,
		layout,
		subtotalPosition: att(outer, 'subtotalTop') === '0' ? 'bottom' : 'top',
		...(second
			? { secondRowField: second, subtotals: subtotal, repeatLabels: repeat[0] ?? false }
			: {
					...(subtotal ? {} : { subtotals: false }),
					...(repeat[0] ? { repeatLabels: true } : {}),
				}),
		valueField: data,
		aggregate,
		valueCaption: decodeEscapes(att(measure, 'name') ?? data),
		fields: names,
		records,
		location,
		partName: part,
		cacheId,
		sourceXml: xml,
		...(col ? { columnField: col } : {}),
	};
	return pivot;
}

/** Admits the supported layout only when its cached cell output agrees with its cache. */
export function readPivotTables(
	workbook: Workbook,
	source: SourceIndex,
	warn: (message: string) => void,
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
	const slicers = [...workbook.source!.parts.keys()].some((part) => /slicer/i.test(part));
	for (const sheet of workbook.sheets)
		if (sheet.partName) {
			const pivots: PivotTable[] = [];
			for (const rel of source.rels(sheet.partName).values()) {
				if (rel.type !== RELATIONSHIP_TYPES.pivotTable) continue;
				const part = source.target(sheet.partName, rel);
				if (!part) continue;
				try {
					const cache = source.targetOfType(part, RELATIONSHIP_TYPES.pivotCacheDefinition);
					if (slicers || (cache && (caches.get(cache) ?? 0) > 1))
						throw new Error('Shared or slicer-connected pivot cache');
					const pivot = readOne(workbook, source, part);
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
