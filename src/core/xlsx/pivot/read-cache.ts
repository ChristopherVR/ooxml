import { RELATIONSHIP_TYPES } from '../../opc/index';
import { NS, children, elements, first, parseXml, relAttr, type XmlElement } from '../../xml/index';
import { parseRange, quoteSheetName, formatRange } from '../address';
import { isErrorCode, type CellValue, type Workbook } from '../model';
import type { SourceIndex } from '../read/package';
import { att, decodeEscapes } from '../read/xml-util';
import { assertPivotAdmission, assertPivotRecords } from './admission';
import { databaseFieldCount } from './date-group-read';
import { isoToSerial } from './dates';
import { pivotSourceRange } from './source';
import { splitCalculatedCacheField } from './calculated';

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

export function readPivotCache(
	workbook: Workbook,
	source: SourceIndex,
	part: string,
	root: XmlElement,
) {
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
	const tableName = att(sourceNode, 'name');
	const sourceTable = tableName && pivotSourceRange(workbook, 0, tableName).tableName;
	const reference =
		sourceTable ||
		(dataSheet && dataRange && workbook.sheets.some((sheet) => sheet.name === dataSheet)
			? `${quoteSheetName(dataSheet)}!${formatRange(dataRange)}`
			: undefined);
	const recordPart = source.targetOfType(cachePart, `${NS.r}/pivotCacheRecords`);
	if (!reference || !recordPart) throw new Error('Unsupported pivot source');
	const recordRoot = parseXml(source.text(recordPart)!).documentElement;
	assertPivotRecords(recordRoot);
	const allFields = children(first(cache, 'cacheFields', NS.x)!, 'cacheField', NS.x);
	const { groupedFields, formula } = splitCalculatedCacheField(allFields);
	// Date groups append derived cache fields that have no records.
	const databaseCount = databaseFieldCount(groupedFields);
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
	return {
		cacheId,
		cache,
		reference,
		recordRoot,
		cacheFieldCount: allFields.length,
		formula,
		groupedFields,
		databaseCount,
		fields,
		shared,
		records,
	};
}
