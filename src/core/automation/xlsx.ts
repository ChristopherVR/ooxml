import { createWorkbook, validateSheetName } from '../xlsx/workbook';
import { loadXlsx } from '../xlsx/read/load';
import { saveXlsx } from '../xlsx/write/save';
import { createEditSession } from '../xlsx/edit/session';
import {
	parseAddress,
	parseRange,
	MAX_COL,
	MAX_ROW,
	formatAddress,
	formatRange,
} from '../xlsx/address';
import { forEachCell, getCell } from '../xlsx/cells';
import { findAll } from '../xlsx/edit/find';
import type { FindQuery } from '../xlsx/edit/types-clipboard';
import { cellView, definedNameList, rangeCells, sheetSummaries } from './xlsx-read';

/**
 * The sheets (index, name, state, used range with its row and column counts, tables), the
 * active sheet, the defined names, the document properties and the loader warnings. Tables per
 * sheet and names are capped at 500 each; `tablesOmitted` and `definedNamesOmitted` count the rest.
 */
export async function inspectXlsx(bytes: Uint8Array) {
	const workbook = await loadXlsx(bytes);
	return {
		sheets: sheetSummaries(workbook),
		activeSheet: workbook.activeSheet,
		...definedNameList(workbook),
		properties: workbook.properties,
		warnings: workbook.warnings,
	};
}

export async function createXlsx(names: readonly string[]): Promise<Uint8Array> {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	if (names.length) {
		const first = names[0]!;
		const problem = validateSheetName(workbook, first, 0);
		if (problem) throw new Error(problem);
		session.renameSheet(0, first);
		for (const name of names.slice(1)) session.addSheet(name);
	}
	return saveXlsx(workbook);
}

/**
 * The cells of an A1 range (at most 10,000), row by row: value, formula, the text as displayed,
 * the number format (absent for General) and, for dates and times, the value in ISO 8601.
 */
export async function readXlsxRange(bytes: Uint8Array, sheetIndex: number, reference: string) {
	const workbook = await loadXlsx(bytes);
	const sheet =
		Number.isSafeInteger(sheetIndex) && sheetIndex >= 0 ? workbook.sheets[sheetIndex] : undefined;
	const range = parseRange(reference);
	if (!sheet || !range || range.end.col > MAX_COL || range.end.row > MAX_ROW)
		throw new Error('Invalid sheet or range');
	return { cells: rangeCells(workbook, sheet, range), warnings: workbook.warnings };
}

export async function setXlsxCells(
	bytes: Uint8Array,
	sheetIndex: number,
	cells: readonly { address: string; input: string }[],
) {
	const workbook = await loadXlsx(bytes);
	if (!Number.isSafeInteger(sheetIndex) || sheetIndex < 0 || !workbook.sheets[sheetIndex])
		throw new Error('Sheet does not exist');
	if (cells.length > 10_000) throw new Error('Edit at most 10000 cells per request');
	const positions = cells.map((cell) => {
		const address = parseAddress(cell.address);
		if (!address) throw new Error(`Invalid cell address: ${cell.address}`);
		return { ...address, input: cell.input };
	});
	const session = createEditSession(workbook);
	session.batch('Automation cell edits', () => {
		for (const cell of positions) session.setCellInput(sheetIndex, cell.row, cell.col, cell.input);
	});
	return { bytes: await saveXlsx(workbook), warnings: workbook.warnings };
}

export interface XlsxFilledCellsOptions {
	/** The most cells returned across all sheets (default 5000). */
	maxCells?: number;
	/** Read only this sheet; every sheet when absent. */
	sheetIndex?: number;
}

/**
 * Every filled cell of every sheet (or of `sheetIndex`), in row-major order, up to `maxCells` in
 * all. Unlike a fixed range it misses nothing inside the cap, and it says how much was left out
 * when the cap is hit.
 */
export async function readXlsxFilledCells(
	bytes: Uint8Array,
	{ maxCells = 5000, sheetIndex }: XlsxFilledCellsOptions = {},
) {
	const workbook = await loadXlsx(bytes);
	if (
		sheetIndex !== undefined &&
		(!Number.isSafeInteger(sheetIndex) || sheetIndex < 0 || !workbook.sheets[sheetIndex])
	)
		throw new Error('Invalid sheet');
	let budget = Math.max(0, Math.floor(maxCells));
	const chosen = workbook.sheets
		.map((sheet, index) => ({ sheet, index }))
		.filter(({ index }) => sheetIndex === undefined || index === sheetIndex);
	const sheets = chosen.map(({ sheet, index }) => {
		const cells: { address: string; value: unknown; formula?: string }[] = [];
		let total = 0;
		forEachCell(sheet, (cell, row, col) => {
			if (cell.value === null && !cell.formula) return;
			total++;
			if (budget <= 0) return;
			budget--;
			cells.push({
				address: formatAddress({ row, col }),
				value: cell.value ?? null,
				...(cell.formula ? { formula: cell.formula } : {}),
			});
		});
		return { index, name: sheet.name, cells, total, omitted: total - cells.length };
	});
	return {
		sheets,
		truncated: sheets.some((s) => s.omitted > 0),
		warnings: workbook.warnings,
	};
}

export interface XlsxFindOptions {
	/** The sheet to search; every sheet when absent. */
	sheetIndex?: number;
	/** An A1 range to search within (needs `sheetIndex`). */
	range?: string;
	/** `formulas` (Excel's default) searches what was typed; `values` the text as displayed. */
	lookIn?: 'formulas' | 'values';
	matchCase?: boolean;
	/** The whole cell content must match. */
	wholeCell?: boolean;
	/** Excel wildcards: `*` any run, `?` one character, `~` escapes. On by default. */
	wildcards?: boolean;
	/** At most this many matches (1 to 1000, default 100); `total` counts them all. */
	maxResults?: number;
}

/**
 * The cells matching `text` the way Excel's Find does, sheet by sheet in row order, each with
 * its sheet, address and the same view as a range read. Bounded by `maxResults`.
 */
export async function findXlsxCells(
	bytes: Uint8Array,
	text: string,
	options: XlsxFindOptions = {},
) {
	if (!text) throw new Error('The search text is empty');
	const max = options.maxResults ?? 100;
	if (!Number.isSafeInteger(max) || max < 1 || max > 1000)
		throw new Error('Return between 1 and 1000 matches per request');
	const workbook = await loadXlsx(bytes);
	const query: FindQuery = { text, lookIn: options.lookIn ?? 'formulas' };
	const { sheetIndex } = options;
	if (sheetIndex !== undefined) {
		if (!Number.isSafeInteger(sheetIndex) || !workbook.sheets[sheetIndex])
			throw new Error('Invalid sheet');
		query.sheet = sheetIndex;
	}
	if (options.range !== undefined) {
		const range = parseRange(options.range);
		if (sheetIndex === undefined) throw new Error('A range needs a sheetIndex');
		if (!range) throw new Error('Invalid range');
		query.range = range;
	}
	if (options.matchCase) query.matchCase = true;
	if (options.wholeCell) query.wholeCell = true;
	if (options.wildcards === false) query.wildcards = false;
	const found = findAll(workbook, query);
	const matches = found.slice(0, max).map((match) => {
		const sheet = workbook.sheets[match.sheet];
		return {
			sheetIndex: match.sheet,
			sheet: sheet?.name ?? '',
			address: formatAddress(match),
			...cellView(workbook, sheet && getCell(sheet, match.row, match.col)),
		};
	});
	return {
		matches,
		total: found.length,
		truncated: found.length > max,
		warnings: workbook.warnings,
	};
}

/**
 * A table found by name (ignoring case, as Excel does) on any sheet: its sheet, range, header and
 * totals rows, column names and cells as a range read returns them (at most 10,000 cells).
 */
export async function readXlsxTable(bytes: Uint8Array, name: string) {
	const workbook = await loadXlsx(bytes);
	const wanted = name.toLowerCase();
	for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
		const table = sheet.tables.find((t) => t.displayName.toLowerCase() === wanted);
		if (!table) continue;
		return {
			sheetIndex,
			sheet: sheet.name,
			name: table.displayName,
			range: formatRange(table.range),
			headerRow: table.headerRow,
			totalsRow: table.totalsRow,
			columns: table.columns.map((column) => column.name),
			cells: rangeCells(workbook, sheet, table.range),
			warnings: workbook.warnings,
		};
	}
	throw new Error(`No table named ${name}`);
}
