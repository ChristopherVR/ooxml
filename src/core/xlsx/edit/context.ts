import { MAX_COL, MAX_ROW, type CellRange, normalizeRange } from '../address.js';
import { getCell, putCell, usedRange } from '../cells.js';
import type { Cell, CellValue, Workbook, Worksheet } from '../model.js';
import { isCellError } from '../model.js';
import { styleAt } from '../styles.js';
import { formatValue, type CalcEngine } from './deps.js';
import type { EditScope } from './history.js';
import type { WorkbookChangeKind } from './types.js';

export interface RunInfo {
	sheet?: number;
	ranges?: CellRange[];
	structural?: boolean;
}

/** What every command module receives from the session. */
export interface EditContext {
	readonly workbook: Workbook;
	readonly calc: CalcEngine;
	/** Base of new sheet names (see `EditSessionOptions.defaultSheetBase`). */
	defaultSheetBase?: string;
	/** Runs `fn` as one undoable step whose effect is confined to `scopes`. */
	run<T>(
		label: string,
		kind: WorkbookChangeKind,
		scopes: EditScope[],
		fn: () => T,
		info?: RunInfo,
	): T;
}

export function sheetAt(workbook: Workbook, index: number): Worksheet {
	const sheet = workbook.sheets[index];
	if (!sheet) throw new RangeError(`No sheet at index ${index}`);
	return sheet;
}

/** The style a new cell at a position inherits: the row's format, else the column's. */
export function baseStyleId(sheet: Worksheet, row: number, col: number): number | undefined {
	const rowStyle = sheet.rowInfo.get(row)?.styleId;
	if (rowStyle) return rowStyle;
	const column = sheet.columns.find((c) => col >= c.min && col <= c.max);
	return column?.styleId || undefined;
}

/** The stored cell at a position, created (with the inherited format) when missing. */
export function ensureCell(sheet: Worksheet, row: number, col: number): Cell {
	const existing = getCell(sheet, row, col);
	if (existing) return existing;
	const cell: Cell = { value: null };
	const styleId = baseStyleId(sheet, row, col);
	if (styleId) cell.styleId = styleId;
	putCell(sheet, row, col, cell);
	return cell;
}

/** Whether a range covers whole columns (every row) or whole rows (every column). */
export const isWholeColumns = (range: CellRange): boolean =>
	range.start.row === 0 && range.end.row >= MAX_ROW;
export const isWholeRows = (range: CellRange): boolean =>
	range.start.col === 0 && range.end.col >= MAX_COL;

/**
 * Clips a range to the sheet's used area so commands over whole rows or columns visit only the
 * stored cells instead of millions of empty positions. Returns undefined when nothing remains.
 */
export function clipToUsed(sheet: Worksheet, range: CellRange): CellRange | undefined {
	const r = normalizeRange(range);
	const used = usedRange(sheet);
	if (!used) return undefined;
	const start = {
		row: Math.max(r.start.row, used.start.row),
		col: Math.max(r.start.col, used.start.col),
	};
	const end = { row: Math.min(r.end.row, used.end.row), col: Math.min(r.end.col, used.end.col) };
	return start.row > end.row || start.col > end.col ? undefined : { start, end };
}

/** Largest number of positions a command visits one by one before clipping to the used area. */
const DENSE_LIMIT = 250_000;

/** Visits every position of a range (clipped to the used area when the range is huge). */
export function forEachPosition(
	sheet: Worksheet,
	range: CellRange,
	visit: (row: number, col: number) => void,
): void {
	let r: CellRange | undefined = normalizeRange(range);
	const size = (r.end.row - r.start.row + 1) * (r.end.col - r.start.col + 1);
	if (size > DENSE_LIMIT) r = clipToUsed(sheet, r);
	if (!r) return;
	for (let row = r.start.row; row <= r.end.row; row++)
		for (let col = r.start.col; col <= r.end.col; col++) visit(row, col);
}

/** The text the grid shows for a cell. */
export function displayText(workbook: Workbook, cell: Cell | undefined): string {
	if (!cell) return '';
	const value = cell.value;
	if (value === null) return '';
	if (isCellError(value)) return value.error;
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (cell.richText && typeof value === 'string') return value;
	const format = styleAt(workbook, cell.styleId).numFmt;
	return formatValue(value, format, { date1904: workbook.date1904 }).text;
}

/** What the user typed for a cell, as the formula bar shows it. */
export function inputText(cell: Cell | undefined): string {
	if (!cell) return '';
	if (cell.formula !== undefined) return `=${cell.formula}`;
	return rawText(cell.value);
}

export function rawText(value: CellValue): string {
	if (value === null) return '';
	if (isCellError(value)) return value.error;
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	return String(value);
}

/** Drops a cell when it carries nothing worth keeping. */
export function pruneCell(sheet: Worksheet, row: number, col: number): void {
	const cell = getCell(sheet, row, col);
	if (!cell) return;
	if (
		(cell.value === null || cell.value === '') &&
		cell.formula === undefined &&
		!cell.styleId &&
		!cell.richText
	) {
		const cells = sheet.rows.get(row);
		cells?.delete(col);
		if (cells && cells.size === 0) sheet.rows.delete(row);
	}
}
