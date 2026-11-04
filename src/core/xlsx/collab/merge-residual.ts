import type * as Y from 'yjs';
import { snapshotEqual } from '../../collab/merge-values';
import { elements, parseXml } from '../../xml/index';
import type { Workbook } from '../model';
import { SourceIndex } from '../read/package';
import { WORKSHEET_CHILD_ORDER } from '../read/worksheet-children';
import { outerXml } from '../read/xml-util';
import { childMap, sharedWorkbook } from './schema';
import { sheetEntries } from './sheets';

const BOOK_MAPPED = new Set([
	'sheets',
	'styles',
	'definedNames',
	'date1904',
	'source',
	'format',
	'warnings',
	'signatures',
]);
const SHEET_MAPPED = new Set([
	'name',
	'sheetId',
	'state',
	'rows',
	'rowInfo',
	'columns',
	'defaultRowHeight',
	'defaultColWidth',
	'merges',
	'tabColor',
	'threadedCommentsXml',
	'partName',
]);
const CELL_MAPPED = new Set([
	'value',
	'formula',
	'arrayRange',
	'styleId',
	'richText',
	'dynamicArray',
	'legacyFormula',
]);
const omit = (object: object, keys: Set<string>) =>
	Object.fromEntries(Object.entries(object).filter(([key]) => !keys.has(key)));

export function workbookResidual(workbook: Workbook, doc: Y.Doc) {
	const entries = sheetEntries(sharedWorkbook(doc).sheets);
	return {
		book: {
			...omit(workbook, BOOK_MAPPED),
			// Producer save stamps are not competing document edits; the remote stamps win.
			properties: omit(workbook.properties, new Set(['modified', 'lastModifiedBy'])),
		},
		sheets: Object.fromEntries(
			workbook.sheets.map((sheet) => {
				const key = entries.find((entry) => entry.sheetId === sheet.sheetId)?.key;
				const cells: Record<string, unknown> = {};
				for (const [row, values] of sheet.rows)
					for (const [col, cell] of values) {
						const extra = omit(cell, CELL_MAPPED);
						if (Object.keys(extra).length) cells[`${row},${col}`] = extra;
					}
				return [key, { properties: omit(sheet, SHEET_MAPPED), cells }];
			}),
		),
	};
}

export function applyWorkbookResidual(
	workbook: Workbook,
	doc: Y.Doc,
	value: ReturnType<typeof workbookResidual>,
) {
	Object.assign(workbook, value.book);
	const entries = sheetEntries(sharedWorkbook(doc).sheets);
	for (const [index, sheet] of workbook.sheets.entries()) {
		const key = entries[index]?.key;
		const residual = key ? value.sheets[key] : undefined;
		if (!residual) continue;
		Object.assign(sheet, residual.properties);
		for (const [address, extra] of Object.entries(residual.cells)) {
			const [row, col] = address.split(',').map(Number);
			const cell =
				row !== undefined && col !== undefined ? sheet.rows.get(row)?.get(col) : undefined;
			if (cell) Object.assign(cell, extra);
		}
	}
}

/** Formula result caches are derived, not user edits. The combined workbook asks Excel to recalc. */
export function stripFormulaCaches(doc: Y.Doc): boolean {
	let formulas = false;
	for (const entry of sheetEntries(sharedWorkbook(doc).sheets)) {
		const cells = childMap(entry.map, 'cells');
		if (!cells) continue;
		for (const [address, raw] of cells) {
			const cell = raw as Record<string, unknown>;
			if (typeof cell?.f !== 'string') continue;
			formulas = true;
			const { v: _value, t: _type, ...rest } = cell;
			cells.set(address, rest);
		}
	}
	return formulas;
}

/** Unknown source fragments of regenerated parts must not silently take the remote branch. */
export function unsupportedSourceChanges(workbooks: Workbook[]): string[] {
	const issues: string[] = [];
	const allowed: Record<string, Set<string>> = {
		worksheet: new Set(WORKSHEET_CHILD_ORDER.filter((name) => name !== 'extLst')),
		workbook: new Set([
			'fileVersion',
			'sheets',
			'definedNames',
			'bookViews',
			'calcPr',
			'workbookPr',
			'workbookProtection',
			'customWorkbookViews',
			'oleSize',
		]),
		styleSheet: new Set([
			'numFmts',
			'fonts',
			'fills',
			'borders',
			'cellStyleXfs',
			'cellXfs',
			'cellStyles',
			'dxfs',
		]),
		sst: new Set(['si']),
	};
	const fingerprints = workbooks.map((book) => {
		const source = book.source && new SourceIndex(book.source.parts);
		const out: Record<string, string[]> = {};
		if (!source) return out;
		for (const [path] of source.parts) {
			const type = source.contentType(path);
			if (!type || !/(?:sheet\.main|worksheet|styles|sharedStrings)\+xml$/u.test(type)) continue;
			const root = parseXml(source.text(path)!).documentElement;
			const owned = allowed[root.localName ?? ''];
			if (!owned) continue;
			const fragments = elements(root)
				.filter((element) => !owned.has(element.localName ?? ''))
				.map(outerXml);
			fragments.push(
				...Array.from(root.attributes)
					.filter((attribute) => attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/')
					.filter(
						(attribute) =>
							root.localName !== 'sst' || !['count', 'uniqueCount'].includes(attribute.name),
					)
					.map(
						(attribute) => `${attribute.namespaceURI ?? ''}|${attribute.name}=${attribute.value}`,
					)
					.sort(),
			);
			if (fragments.length) out[path] = fragments;
		}
		return out;
	});
	for (const key of new Set(fingerprints.flatMap(Object.keys)))
		if (
			!snapshotEqual(fingerprints[0]?.[key], fingerprints[1]?.[key]) ||
			!snapshotEqual(fingerprints[0]?.[key], fingerprints[2]?.[key])
		)
			issues.push(`Unmapped source XML changed in ${key}.`);
	return issues;
}
