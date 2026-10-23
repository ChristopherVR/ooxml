// Reads one worksheet's rows without building the workbook: for imports and agents that walk a
// large sheet once. The workbook, its relationships, styles, shared strings and cell metadata are
// read as `loadXlsx` reads them; the sheet part is inflated and read in pieces (see rows-stream),
// so what is held is those parts, the package bytes the caller passed in and the current piece
// of the sheet, never a `Workbook`.
import type JSZip from 'jszip';
import {
	RELATIONSHIP_TYPES,
	parseRelationships,
	relationshipsPartFor,
	resolvePartPath,
	type Relationship,
} from '../../opc/index';
import type { Cell, CellValue } from '../model';
import { type NumberFormatView, numberFormatView } from '../numfmt/view';
import { resolveLoadLimits, type XlsxLoadLimits } from './limits';
import { parseDynamicArrayMetadata } from './metadata';
import { openPackage } from './package';
import { sheetRows } from './rows-stream';
import { parseSharedStrings } from './shared-strings';
import type { CellContext } from './sheet-data';
import { parseStyles } from './styles';
import { parseWorkbookPart } from './workbook-part';

export interface ReadXlsxRowsOptions {
	/** The sheet to read: its position in the workbook (0 is the first) or its name. Default 0. */
	sheet?: number | string;
	/**
	 * Size limits, as for `loadXlsx`: declared sizes are checked when the package is opened and
	 * the parts that are read count their bytes as they inflate.
	 */
	limits?: XlsxLoadLimits;
	/** Called once for each distinct problem the reader reports (a `loadXlsx` warning). */
	onWarning?: (message: string) => void;
}

export interface XlsxRowCell extends NumberFormatView {
	/** Zero-based column index. */
	col: number;
	/** The value as `loadXlsx` reads it: dates are serial numbers (see `iso`). */
	value: CellValue;
	/** The formula without its `=` (shared formulas expanded); `value` is its cached result. */
	formula?: string;
}

export interface XlsxRow {
	/** Zero-based row index. */
	row: number;
	/** The row's cells with a value or a formula, by column. */
	cells: XlsxRowCell[];
}

/**
 * The rows of one worksheet in the order the file stores them (ascending in files Excel writes),
 * each with the cells that hold a value or a formula; rows without one are left out. Values,
 * formulas, dates and number formats are what `loadXlsx` and the automation reads give for the
 * same cells. Breaking out of the loop stops inflating the sheet. A cell stored under another
 * row's element comes as a row of its own.
 */
export async function* readXlsxRows(
	input: Uint8Array | ArrayBuffer,
	options: ReadXlsxRowsOptions = {},
): AsyncGenerator<XlsxRow, void, undefined> {
	const limits = resolveLoadLimits(options.limits);
	const zip = await openPackage(input, limits);
	const warned = new Set<string>();
	const warn = (message: string) => {
		if (warned.has(message)) return;
		warned.add(message);
		options.onWarning?.(message);
	};
	const parts = new PartReader(zip);
	const rootRels = parseRelationships(await parts.text('_rels/.rels'));
	const fromRels = targetOfType('', rootRels, RELATIONSHIP_TYPES.officeDocument);
	const bookPart = fromRels && parts.has(fromRels) ? fromRels : 'xl/workbook.xml';
	const bookXml = await parts.text(bookPart);
	if (bookXml === undefined) throw new Error('XLSX package has no workbook part (xl/workbook.xml)');
	const book = parseWorkbookPart(bookXml);
	const rels = parseRelationships(await parts.text(relationshipsPartFor(bookPart)));
	const entry = pickSheet(book.sheets, options.sheet ?? 0);
	const rel = rels.get(entry.relId);
	const target = rel && rel.mode !== 'External' ? resolvePartPath(bookPart, rel.target) : undefined;
	if (rel?.type !== RELATIONSHIP_TYPES.worksheet || !target) {
		const kind = rel?.type.split('/').pop();
		throw new Error(
			`Sheet "${entry.name}" ${kind ? `is a ${kind}` : 'has no part'}; it has no rows to read.`,
		);
	}
	const partOf = (type: string) => targetOfType(bookPart, rels, type);
	const warnings: string[] = [];
	const styles = parseStyles(await parts.text(partOf(RELATIONSHIP_TYPES.styles)), warnings);
	for (const message of warnings) warn(message);
	const ctx: CellContext = {
		sharedStrings: parseSharedStrings(
			await parts.text(partOf(RELATIONSHIP_TYPES.sharedStrings)),
			styles.palette,
			limits.maxXmlElements,
		),
		xfMap: styles.xfMap,
		date1904: book.date1904,
		palette: styles.palette,
		dynamicCells: parseDynamicArrayMetadata(
			await parts.text(partOf(RELATIONSHIP_TYPES.sheetMetadata)),
		),
		warn,
	};
	const sheet = parts.entry(target);
	if (!sheet) {
		warn(`Sheet "${entry.name}" has no part ${target}; it was loaded empty.`);
		return;
	}
	const formats = styles.styles.map((style) => style.numFmt);
	const cellOf = (col: number, cell: Cell): XlsxRowCell => ({
		col,
		value: cell.value,
		...(cell.formula ? { formula: cell.formula } : {}),
		...numberFormatView(
			formats[cell.styleId ?? 0] ?? formats[0] ?? 'General',
			cell.value,
			book.date1904,
		),
	});
	for await (const [row, cells] of sheetRows(ctx, sheet, parts.read, limits)) {
		const out: XlsxRowCell[] = [];
		for (const [col, cell] of cells)
			if (cell.value !== null || cell.formula) out.push(cellOf(col, cell));
		if (out.length) yield { row, cells: out.sort((a, b) => a.col - b.col) };
	}
}

/** The sheet at `wanted` (a position) or named `wanted` (ignoring case, as Excel does). */
function pickSheet<T extends { name: string }>(sheets: readonly T[], wanted: number | string): T {
	const found =
		typeof wanted === 'number'
			? Number.isSafeInteger(wanted) && wanted >= 0
				? sheets[wanted]
				: undefined
			: sheets.find((sheet) => sheet.name.toLowerCase() === wanted.toLowerCase());
	if (!found)
		throw new RangeError(
			typeof wanted === 'number'
				? `The workbook has no sheet at position ${wanted}`
				: `The workbook has no sheet named "${wanted}"`,
		);
	return found;
}

function targetOfType(
	partName: string,
	rels: ReadonlyMap<string, Relationship>,
	type: string,
): string | undefined {
	for (const rel of rels.values())
		if (rel.type === type)
			return rel.mode === 'External'
				? undefined
				: resolvePartPath(partName === '' ? '/' : partName, rel.target);
	return undefined;
}

/** The package's parts by name (ignoring case, as OPC does), inflated only when asked for. */
class PartReader {
	private readonly entries = new Map<string, JSZip.JSZipObject>();
	/** Bytes inflated so far, for the total limit of the part read in pieces. */
	read = 0;
	private readonly decoder = new TextDecoder('utf-8');

	constructor(zip: JSZip) {
		for (const entry of Object.values(zip.files))
			if (!entry.dir) this.entries.set(entry.name.replace(/^\//, '').toLowerCase(), entry);
	}

	entry(partName: string): JSZip.JSZipObject | undefined {
		return this.entries.get(partName.toLowerCase());
	}

	has(partName: string): boolean {
		return this.entries.has(partName.toLowerCase());
	}

	/** A part decoded as UTF-8 (BOM stripped), or undefined when absent. */
	async text(partName: string | undefined): Promise<string | undefined> {
		const entry = partName ? this.entry(partName) : undefined;
		if (!entry) return undefined;
		const bytes = await entry.async('uint8array');
		this.read += bytes.byteLength;
		return this.decoder.decode(bytes);
	}
}
