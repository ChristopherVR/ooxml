/**
 * Writes chart-edited values back into the workbook a chart embeds (`c:externalData` pointing at
 * an `.xlsx` package part), so the product's "Edit Data in Excel" shows what the chart shows.
 *
 * A chart caches its data twice: in the chart part's `c:numCache`/`c:strCache`, and in the
 * workbook. Each series' `c:tx`, `c:cat` and `c:val` (or `c:xVal`/`c:yVal`) carries a `c:f`
 * formula naming the cells it was read from (`Sheet1!$B$2:$B$4`). This module resolves that
 * formula to a worksheet part and the cells it names, and rewrites only those cells; other
 * sheets, styles, formatting and unrelated cells are left as they are.
 *
 * Scope: a straight vector reference (one row or one column) of `cellCount` cells is written;
 * anything else (a structural change the formula was not updated for, a defined name, a
 * reference this module does not parse, a missing worksheet) is counted in
 * {@link ChartWorkbookWriteResult.unresolved} for the caller to report. The non-atomic mode never
 * throws. Writing a value into a cell drops a formula the cell held, as typing over it in Excel
 * does.
 *
 * Moved from the PowerPoint area (`pptx/core/utils/chart-xlsx-writer.ts`) and ported from the
 * `fast-xml-parser` object tree to the shared DOM (`xml` area).
 *
 * @module chart/workbook/writer
 */
import type JSZip from 'jszip';
import { openZip } from '../../opc/zip-limits';
import { parseRelationships, resolvePartPath } from '../../opc/relationships';
import { buildXml, parseXml, type XmlDocument } from '../../xml/index';
import { parseCellAddress, parseChartFormulaRange, rangeCellAddress } from './cellref';
import type { ChartFormulaRange } from './cellref';
import {
	childNamed,
	childrenNamed,
	restoreWorksheetCells,
	worksheetData,
	writeWorksheetCells,
	type ChartCellChange,
	type ChartCellWrite,
} from './sheet-cells';
import { sharedStringTexts, syncWorksheetTables, worksheetCellText } from './tables';
import { validateAtomicChartWorkbookWrites } from './validation';

/** One cell of a chart data reference: its position in the range and its new text. */
export interface ChartWorkbookCell {
	index: number;
	/** The cell's new content; empty clears it. */
	value: string;
}

/** One chart data reference (series name, categories or values) to write back. */
export interface ChartWorkbookWrite {
	/** The `c:f` formula naming the source cells (e.g. `"Sheet1!$B$2:$B$4"`). */
	formula: string;
	isNumeric: boolean;
	/** The cells to change; every other cell of the range stays as it is. */
	cells: readonly ChartWorkbookCell[];
	/** How many cells the formula names; a range of another size is not written. */
	cellCount: number;
	/**
	 * The range holds nothing but `cells`: a cell of it that the chart would read as a point the
	 * write does not list is emptied (a number in a value range, anything in a text range). For a
	 * range the chart did not read before, whose cells may hold anything.
	 */
	only?: boolean;
}

/** Every value of a whole range as cells to write. */
export const chartWorkbookCells = (values: readonly string[]): ChartWorkbookCell[] =>
	values.map((value, index) => ({ index, value }));

/** A worksheet cell a write-back changed, to put back with {@link restoreChartWorkbookCells}. */
export interface ChartWorkbookCellChange extends ChartCellChange {
	sheetPath: string;
}

/** Result of applying a batch of writes to an embedded workbook. */
export interface ChartWorkbookWriteResult {
	/** Updated xlsx bytes; present only when at least one cell changed. */
	bytes?: Uint8Array;
	/** Writes that could not be matched to worksheet cells (see the module comment). */
	unresolved: number;
	/** The cells the writes changed, with their markup before and after. */
	changes?: ChartWorkbookCellChange[];
}

export interface ChartWorkbookWriteOptions {
	/** Throw instead of writing anything when any write cannot be applied exactly. */
	atomic?: boolean;
	/** With `atomic`: the cells must still hold the values the chart cache had. */
	requireCurrentValues?: boolean;
	/** Check every write as it would be applied, but return no bytes (nothing is packaged). */
	dryRun?: boolean;
}

const WORKBOOK_PART = 'xl/workbook.xml';
const WORKBOOK_RELS = 'xl/_rels/workbook.xml.rels';

/** Opens an embedded package with the shared default zip limits. */
export const openEmbeddedPackage = (bytes: Uint8Array): Promise<JSZip> =>
	openZip(bytes, { label: 'Embedded workbook' });

/** `sheet name (lower case) -> worksheet part` from `xl/workbook.xml` and its relationships. */
export async function worksheetPathsByName(zip: JSZip): Promise<Map<string, string>> {
	const result = new Map<string, string>();
	const workbookXml = await zip.file(WORKBOOK_PART)?.async('string');
	const relsXml = await zip.file(WORKBOOK_RELS)?.async('string');
	if (!workbookXml || !relsXml) return result;
	try {
		const rels = parseRelationships(relsXml);
		const root = parseXml(workbookXml, { label: 'Workbook' }).documentElement ?? undefined;
		for (const sheets of childrenNamed(root, 'sheets'))
			for (const sheet of childrenNamed(sheets, 'sheet')) {
				const name = sheet.getAttribute('name');
				const relAttr = Array.from(sheet.attributes).find((item) => item.localName === 'id');
				const rel = relAttr ? rels.get(relAttr.value) : undefined;
				if (name && rel && rel.mode !== 'External')
					result.set(name.toLowerCase(), resolvePartPath(WORKBOOK_PART, rel.target));
			}
	} catch {
		// A malformed workbook part: return what was found, if anything.
	}
	return result;
}

interface ResolvedWrite {
	address: (index: number) => string;
	cells: readonly ChartWorkbookCell[];
	isNumeric: boolean;
	/** With `only`: the range and its size, to find the cells it does not list. */
	only?: { range: ChartFormulaRange; count: number };
}

/** Empties the cells of an `only` write's range that would read as a point it does not list. */
function unlistedPointCells(doc: XmlDocument, write: ResolvedWrite): ChartCellWrite[] {
	if (!write.only) return [];
	const { range, count } = write.only;
	const listed = new Set(write.cells.map(({ index }) => write.address(index)));
	const clears: ChartCellWrite[] = [];
	for (const row of childrenNamed(worksheetData(doc), 'row'))
		for (const cell of childrenNamed(row, 'c')) {
			const ref = cell.getAttribute('r') ?? '';
			const at = parseCellAddress(ref);
			if (!at || listed.has(ref)) continue;
			const index =
				range.start.col === range.end.col && at.col === range.start.col
					? at.row - range.start.row
					: range.start.row === range.end.row && at.row === range.start.row
						? at.col - range.start.col
						: -1;
			if (index < 0 || index >= count) continue;
			const type = cell.getAttribute('t') ?? 'n';
			const holds = childNamed(cell, 'v') !== undefined || childNamed(cell, 'is') !== undefined;
			// A text cell in a value range is no point of it, and keeps its text.
			if (holds && (!write.isNumeric || type === 'n'))
				clears.push({ ref, isNumeric: write.isNumeric, value: '' });
		}
	return clears;
}

/**
 * Rewrites an embedded workbook's cells to match a chart's edited data. Without `atomic` it never
 * throws: a corrupt package, a missing worksheet or a formula it cannot resolve is counted in
 * `unresolved`, so the caller keeps the chart's cache-only edit and reports it.
 */
export async function writeChartWorkbookUpdates(
	xlsxBytes: Uint8Array,
	writes: readonly ChartWorkbookWrite[],
	options: ChartWorkbookWriteOptions = {},
): Promise<ChartWorkbookWriteResult> {
	if (writes.length === 0) return { unresolved: 0 };
	let zip: JSZip;
	try {
		zip = await openEmbeddedPackage(xlsxBytes);
	} catch {
		if (options.atomic)
			throw new Error('Chart label replacement requires a readable embedded workbook.');
		return { unresolved: writes.length };
	}
	const sheetPathByName = await worksheetPathsByName(zip);
	if (options.atomic)
		await validateAtomicChartWorkbookWrites(
			zip,
			sheetPathByName,
			writes,
			options.requireCurrentValues ?? false,
		);
	const writesBySheet = new Map<string, ResolvedWrite[]>();
	let unresolved = 0;
	for (const write of writes) {
		const range = parseChartFormulaRange(write.formula);
		const sheetPath = range ? sheetPathByName.get(range.sheetName.toLowerCase()) : undefined;
		const address = range ? rangeCellAddress(range, write.cellCount) : undefined;
		if (!range || !sheetPath || !address) {
			unresolved += 1;
			continue;
		}
		const bucket = writesBySheet.get(sheetPath) ?? [];
		bucket.push({
			address,
			cells: write.cells,
			isNumeric: write.isNumeric,
			...(write.only ? { only: { range, count: write.cellCount } } : {}),
		});
		writesBySheet.set(sheetPath, bucket);
	}
	const changes: ChartWorkbookCellChange[] = [];
	const shared = await sharedStringTexts(zip).catch(() => []);
	for (const [sheetPath, resolved] of writesBySheet) {
		const file = zip.file(sheetPath);
		if (!file) {
			unresolved += resolved.length;
			continue;
		}
		try {
			const doc = parseXml(await file.async('string'), { label: 'Worksheet' });
			const cellWrites: ChartCellWrite[] = resolved.flatMap((write) => [
				...write.cells.map(({ index, value }) => ({
					ref: write.address(index),
					isNumeric: write.isNumeric,
					value,
				})),
				...unlistedPointCells(doc, write),
			]);
			const written = writeWorksheetCells(doc, cellWrites, shared);
			if (written.length > 0) {
				await syncWorksheetTables(zip, sheetPath, doc, cellWrites, shared);
				if (!options.dryRun) zip.file(sheetPath, buildXml(doc));
				for (const change of written) changes.push({ sheetPath, ...change });
			}
		} catch {
			unresolved += resolved.length;
		}
	}
	if (options.atomic && unresolved > 0)
		throw new Error('Chart label replacement could not update every source range.');
	if (changes.length === 0 || options.dryRun) return { unresolved };
	return { bytes: await packed(zip), unresolved, changes };
}

const packed = (zip: JSZip) =>
	zip.generateAsync({
		type: 'uint8array',
		compression: 'DEFLATE',
		compressionOptions: { level: 6 },
	});

/**
 * Puts back the cells a write-back changed (`ChartWorkbookWriteResult.changes`) wherever they still
 * hold what it wrote, leaving every cell written since as it is. Returns the updated bytes, or
 * `undefined` when no cell was put back.
 */
export async function restoreChartWorkbookCells(
	xlsxBytes: Uint8Array,
	changes: readonly ChartWorkbookCellChange[],
): Promise<Uint8Array | undefined> {
	if (changes.length === 0) return undefined;
	const zip = await openEmbeddedPackage(xlsxBytes);
	const shared = await sharedStringTexts(zip).catch(() => []);
	let restored = false;
	for (const sheetPath of new Set(changes.map((change) => change.sheetPath))) {
		const file = zip.file(sheetPath);
		if (!file) continue;
		const doc = parseXml(await file.async('string'), { label: 'Worksheet' });
		const own = changes.filter((change) => change.sheetPath === sheetPath);
		const refs = new Set(restoreWorksheetCells(doc, own));
		if (refs.size === 0) continue;
		// A restored header names its table column again, as a written one did.
		const cells = own
			.filter((change) => refs.has(change.ref) && change.before !== undefined)
			.map(({ ref }) => ({ ref, isNumeric: false, value: worksheetCellText(doc, ref, shared) }));
		await syncWorksheetTables(zip, sheetPath, doc, cells, shared);
		zip.file(sheetPath, buildXml(doc));
		restored = true;
	}
	return restored ? packed(zip) : undefined;
}
