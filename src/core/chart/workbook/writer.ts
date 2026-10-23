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
 * Scope: a straight vector reference (one row or one column) whose length matches the values is
 * written; anything else (a structural change the formula was not updated for, a defined name, a
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
import { buildXml, parseXml } from '../../xml/index';
import { expandRangeAddresses, parseChartFormulaRange } from './cellref';
import { applyCellWritesToWorksheet, childrenNamed, type ChartCellWrite } from './sheet-cells';
import { validateAtomicChartWorkbookWrites } from './validation';

/** One chart data reference (series name, categories or values) to write back. */
export interface ChartWorkbookWrite {
	/** The `c:f` formula naming the source cells (e.g. `"Sheet1!$B$2:$B$4"`). */
	formula: string;
	isNumeric: boolean;
	/** Values in on-sheet order, one per cell the formula range names. */
	values: string[];
}

/** Result of applying a batch of writes to an embedded workbook. */
export interface ChartWorkbookWriteResult {
	/** Updated xlsx bytes; present only when at least one cell changed. */
	bytes?: Uint8Array;
	/** Writes that could not be matched to worksheet cells (see the module comment). */
	unresolved: number;
}

export interface ChartWorkbookWriteOptions {
	/** Throw instead of writing anything when any write cannot be applied exactly. */
	atomic?: boolean;
	/** With `atomic`: the cells must still hold the values the chart cache had. */
	requireCurrentValues?: boolean;
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
	refs: string[];
	values: string[];
	isNumeric: boolean;
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
		const refs = range ? expandRangeAddresses(range, write.values.length) : undefined;
		if (!range || !sheetPath || !refs) {
			unresolved += 1;
			continue;
		}
		const bucket = writesBySheet.get(sheetPath) ?? [];
		bucket.push({ refs, values: write.values, isNumeric: write.isNumeric });
		writesBySheet.set(sheetPath, bucket);
	}
	let changed = false;
	for (const [sheetPath, resolved] of writesBySheet) {
		const file = zip.file(sheetPath);
		if (!file) {
			unresolved += resolved.length;
			continue;
		}
		try {
			const doc = parseXml(await file.async('string'), { label: 'Worksheet' });
			const cellWrites: ChartCellWrite[] = resolved.flatMap(({ refs, values, isNumeric }) =>
				refs.map((ref, index) => ({ ref, isNumeric, value: values[index] ?? '' })),
			);
			if (applyCellWritesToWorksheet(doc, cellWrites)) {
				zip.file(sheetPath, buildXml(doc));
				changed = true;
			}
		} catch {
			unresolved += resolved.length;
		}
	}
	if (options.atomic && unresolved > 0)
		throw new Error('Chart label replacement could not update every source range.');
	if (!changed) return { unresolved };
	const bytes = await zip.generateAsync({
		type: 'uint8array',
		compression: 'DEFLATE',
		compressionOptions: { level: 6 },
	});
	return { bytes, unresolved };
}
