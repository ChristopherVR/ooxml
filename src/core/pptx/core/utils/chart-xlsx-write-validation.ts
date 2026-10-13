import type { XMLParser } from 'fast-xml-parser';
import type JSZip from 'jszip';

import type { XmlObject } from '../types';
import { expandRangeAddresses, parseChartFormulaRange } from './chart-xlsx-cellref';
import type { PptxChartWorkbookWrite } from './chart-xlsx-writer';
import { xmlAttr, xmlChild, xmlChildren, xmlText } from './xml-access';

function unsupported(reason: string): never {
	throw new Error(`Chart label replacement requires direct literal workbook cells: ${reason}.`);
}

/** Validate the complete transaction before the existing writer changes any worksheet. */
export async function validateAtomicChartWorkbookWrites(
	zip: JSZip,
	parser: XMLParser,
	paths: ReadonlyMap<string, string>,
	writes: readonly PptxChartWorkbookWrite[],
	requireCurrentValues = false,
): Promise<void> {
	const relationships = xmlChild(
		parser.parse(
			(await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? '',
		) as XmlObject,
		'Relationships',
	);
	for (const rel of xmlChildren(relationships, 'Relationship')) {
		if (xmlAttr(rel, 'TargetMode') === 'External') {
			unsupported('external worksheet relationships are not supported');
		}
	}
	const sheets = new Map<string, XmlObject>();
	const cellsByPath = new Map<string, Map<string, XmlObject>>();
	for (const path of new Set(paths.values())) {
		const file = zip.file(path);
		if (!file) {
			unsupported('a referenced worksheet is missing');
		}
		let tree: XmlObject;
		try {
			tree = parser.parse(await file.async('string')) as XmlObject;
		} catch {
			unsupported('a worksheet cannot be parsed');
		}
		const data = xmlChild(xmlChild(tree, 'worksheet'), 'sheetData');
		if (!data) {
			unsupported('a worksheet has no cell data');
		}
		// The existing cell writer emits unprefixed row/cell nodes. Do not create a
		// second set alongside prefixed source nodes until that path supports them.
		if (Object.keys(data).some((key) => key.includes(':'))) {
			unsupported('prefixed worksheet cell nodes are not supported');
		}
		const cells = new Map<string, XmlObject>();
		for (const row of xmlChildren(data, 'row')) {
			if (Object.keys(row).some((key) => key.includes(':'))) {
				unsupported('prefixed worksheet cell nodes are not supported');
			}
			for (const cell of xmlChildren(row, 'c')) {
				const ref = xmlAttr(cell, 'r');
				if (!ref || cells.has(ref)) {
					unsupported('duplicate or unidentified worksheet cells are not supported');
				}
				cells.set(ref, cell);
				if (Object.keys(cell).some((key) => key === 'f' || key.endsWith(':f'))) {
					unsupported('formula-dependent workbooks are not supported');
				}
			}
		}
		sheets.set(path, data);
		cellsByPath.set(path, cells);
	}
	const sharedFile = zip.file('xl/sharedStrings.xml');
	const shared = sharedFile
		? xmlChildren(
				xmlChild(parser.parse(await sharedFile.async('string')) as XmlObject, 'sst'),
				'si',
			)
		: [];
	const destinations = new Map<string, { value: string; numeric: boolean }>();
	for (const write of writes) {
		const range = parseChartFormulaRange(write.formula);
		if (!range || /[\[\]:!]/u.test(range.sheetName)) {
			unsupported('a range is not a direct local reference');
		}
		if (
			[range.start, range.end].some(
				(cell) => cell.row < 0 || cell.row >= 1048576 || cell.col < 0 || cell.col >= 16384,
			)
		) {
			unsupported('a source cell lies outside the worksheet limits');
		}
		const addresses = expandRangeAddresses(range, write.values.length);
		const path = paths.get(range.sheetName.toLowerCase());
		const data = path ? sheets.get(path) : undefined;
		if (!addresses || !path || !data) {
			unsupported('a source range cannot be resolved');
		}
		for (const [index, address] of addresses.entries()) {
			const value = write.values[index];
			const key = `${path}!${address}`;
			const prior = destinations.get(key);
			if (prior && (prior.value !== value || prior.numeric !== write.isNumeric)) {
				unsupported('overlapping source ranges request conflicting values');
			}
			destinations.set(key, { value, numeric: write.isNumeric });
			const cell = cellsByPath.get(path)?.get(address);
			if (!cell) {
				unsupported('a source cell is missing');
			}
			const type = xmlAttr(cell, 't') ?? 'n';
			const text = type === 's' ? shared[Number(xmlText(cell.v))] : xmlChild(cell, 'is');
			if (
				text &&
				Object.keys(text).some((node) => !node.startsWith('@_') && node !== 't' && node !== '#text')
			) {
				unsupported('rich worksheet strings require source-workbook editing');
			}
			if (requireCurrentValues) {
				const current = text ? xmlText(text.t) : xmlText(cell.v);
				const equal = write.isNumeric
					? current !== undefined && Number(current) === Number(value)
					: current === value;
				if (!equal) {
					unsupported('the workbook and chart caches disagree');
				}
			}
			if (write.isNumeric && !Number.isFinite(Number(value))) {
				unsupported('a numeric cache value is not finite');
			}
			if (!['n', 's', 'inlineStr', 'str'].includes(type)) {
				unsupported('a source cell has an unsupported value type');
			}
		}
	}
}
