/**
 * The all-or-nothing check an atomic workbook write runs before any worksheet changes: every
 * reference resolves to existing literal cells of a local worksheet, no write conflicts with
 * another, and (optionally) the cells still hold what the chart cache says.
 *
 * @module chart/workbook/validation
 */
import type JSZip from 'jszip';
import { parseXml, type XmlElement } from '../../xml/index';
import { expandRangeAddresses, parseChartFormulaRange } from './cellref';
import { childNamed, childrenNamed, worksheetData } from './sheet-cells';
import type { ChartWorkbookWrite } from './writer';

function unsupported(reason: string): never {
	throw new Error(`Chart label replacement requires direct literal workbook cells: ${reason}.`);
}

async function readCells(zip: JSZip, path: string): Promise<Map<string, XmlElement>> {
	const file = zip.file(path);
	if (!file) unsupported('a referenced worksheet is missing');
	let data: XmlElement | undefined;
	try {
		data = worksheetData(parseXml(await file.async('string'), { label: 'Worksheet' }));
	} catch {
		unsupported('a worksheet cannot be parsed');
	}
	if (!data) unsupported('a worksheet has no cell data');
	const cells = new Map<string, XmlElement>();
	for (const row of childrenNamed(data, 'row'))
		for (const cell of childrenNamed(row, 'c')) {
			const ref = cell.getAttribute('r');
			if (!ref || cells.has(ref))
				unsupported('duplicate or unidentified worksheet cells are not supported');
			cells.set(ref, cell);
			if (childNamed(cell, 'f')) unsupported('formula-dependent workbooks are not supported');
		}
	return cells;
}

/** The text element a string cell reads from: its shared string item or its inline string. */
function stringItem(cell: XmlElement, shared: readonly XmlElement[]): XmlElement | undefined {
	const type = cell.getAttribute('t') || 'n';
	if (type === 's') return shared[Number(childNamed(cell, 'v')?.textContent)];
	return childNamed(cell, 'is');
}

/** Validate the complete transaction before the writer changes any worksheet. */
export async function validateAtomicChartWorkbookWrites(
	zip: JSZip,
	paths: ReadonlyMap<string, string>,
	writes: readonly ChartWorkbookWrite[],
	requireCurrentValues = false,
): Promise<void> {
	const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
	if (relsXml)
		for (const rel of childrenNamed(parseXml(relsXml).documentElement ?? undefined, 'Relationship'))
			if (rel.getAttribute('TargetMode') === 'External')
				unsupported('external worksheet relationships are not supported');
	const cellsByPath = new Map<string, Map<string, XmlElement>>();
	for (const path of new Set(paths.values())) cellsByPath.set(path, await readCells(zip, path));
	const sharedXml = await zip.file('xl/sharedStrings.xml')?.async('string');
	const shared = sharedXml
		? childrenNamed(parseXml(sharedXml).documentElement ?? undefined, 'si')
		: [];
	const destinations = new Map<string, { value: string; numeric: boolean }>();
	for (const write of writes) {
		const range = parseChartFormulaRange(write.formula);
		if (!range || /[[\]:!]/u.test(range.sheetName))
			unsupported('a range is not a direct local reference');
		if (
			[range.start, range.end].some(
				(cell) => cell.row < 0 || cell.row >= 1048576 || cell.col < 0 || cell.col >= 16384,
			)
		)
			unsupported('a source cell lies outside the worksheet limits');
		const addresses = expandRangeAddresses(range, write.values.length);
		const path = paths.get(range.sheetName.toLowerCase());
		const cells = path ? cellsByPath.get(path) : undefined;
		if (!addresses || !path || !cells) unsupported('a source range cannot be resolved');
		for (const [index, address] of addresses.entries()) {
			const value = write.values[index] ?? '';
			const key = `${path}!${address}`;
			const prior = destinations.get(key);
			if (prior && (prior.value !== value || prior.numeric !== write.isNumeric))
				unsupported('overlapping source ranges request conflicting values');
			destinations.set(key, { value, numeric: write.isNumeric });
			const cell = cells.get(address);
			if (!cell) unsupported('a source cell is missing');
			const type = cell.getAttribute('t') || 'n';
			const text = stringItem(cell, shared);
			if (
				text &&
				Array.from(text.childNodes).some(
					(node) => node.nodeType === 1 && (node as XmlElement).localName !== 't',
				)
			)
				unsupported('rich worksheet strings require source-workbook editing');
			if (requireCurrentValues) {
				const current = text
					? (childNamed(text, 't')?.textContent ?? undefined)
					: (childNamed(cell, 'v')?.textContent ?? undefined);
				const equal = write.isNumeric
					? current !== undefined && Number(current) === Number(value)
					: current === value;
				if (!equal) unsupported('the workbook and chart caches disagree');
			}
			if (write.isNumeric && !Number.isFinite(Number(value)))
				unsupported('a numeric cache value is not finite');
			if (!['n', 's', 'inlineStr', 'str'].includes(type))
				unsupported('a source cell has an unsupported value type');
		}
	}
}
