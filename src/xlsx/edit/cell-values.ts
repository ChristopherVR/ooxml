import {
	type CellAddress,
	type CellRange,
	formatAddress,
	normalizeRange,
	rangesIntersect,
} from '../address.js';
import { deleteCell, forEachCellInRange, getCell } from '../cells.js';
import type { Cell, CellValue, Worksheet } from '../model.js';
import { applyStylePatch, styleAt } from '../styles.js';
import {
	type EditContext,
	baseStyleId,
	displayText,
	ensureCell,
	pruneCell,
	sheetAt,
} from './context.js';
import { parseCellInput } from './deps.js';
import type { EditScope, SheetPart } from './history.js';
import { cellRange, subtractRange } from './range-math.js';
import type { ClearWhat } from './types.js';

/** The table whose header row holds a position, if any. */
function tableHeaderAt(sheet: Worksheet, row: number, col: number) {
	return sheet.tables.find(
		(t) =>
			t.headerRow &&
			t.range.start.row === row &&
			col >= t.range.start.col &&
			col <= t.range.end.col,
	);
}

/** Keeps a table's column name in step with its header cell. */
export function syncTableHeader(
	ctx: EditContext,
	sheet: Worksheet,
	row: number,
	col: number,
): void {
	const table = tableHeaderAt(sheet, row, col);
	if (!table) return;
	const column = table.columns[col - table.range.start.col];
	if (!column) return;
	const text = displayText(ctx.workbook, getCell(sheet, row, col)).trim();
	if (text) column.name = text;
}

/**
 * Undo scopes for writing to cells: the cells, plus the table list when a table header changes
 * (the column names follow the header text). `parts` adds other sheet properties the edit changes.
 */
export function writeScopes(
	sheet: Worksheet,
	index: number,
	range: CellRange,
	parts: SheetPart[] = [],
): EditScope[] {
	const r = normalizeRange(range);
	const touchesHeader = sheet.tables.some(
		(t) =>
			t.headerRow &&
			rangesIntersect(
				r,
				cellRange(t.range.start.row, t.range.start.col, t.range.start.row, t.range.end.col),
			),
	);
	const all = touchesHeader && !parts.includes('tables') ? [...parts, 'tables' as const] : parts;
	const scopes: EditScope[] = [{ kind: 'cells', sheet: index, ranges: [r] }];
	if (all.length) scopes.push({ kind: 'parts', sheet: index, parts: all });
	return scopes;
}

/** The single undo scope of {@link writeScopes}, kept for callers that want one. */
export function writeScope(sheet: Worksheet, index: number, range: CellRange): EditScope {
	const scopes = writeScopes(sheet, index, range);
	return scopes.length === 1 && scopes[0] ? scopes[0] : { kind: 'sheet', sheet: index };
}

/** Writes typed text into a cell without opening an undo step (shared by paste and replace). */
export function writeInput(
	ctx: EditContext,
	sheet: Worksheet,
	row: number,
	col: number,
	text: string,
): void {
	const { workbook } = ctx;
	const existing = getCell(sheet, row, col);
	const format = styleAt(workbook, existing?.styleId ?? baseStyleId(sheet, row, col)).numFmt;
	if (text === '') {
		if (existing) clearContents(existing);
		pruneCell(sheet, row, col);
		return;
	}
	const cell = ensureCell(sheet, row, col);
	clearContents(cell);
	if (format === '@') {
		cell.value = text;
		return;
	}
	const parsed = parseCellInput(text, { date1904: workbook.date1904 });
	if (parsed.formula !== undefined) cell.formula = parsed.formula;
	else cell.value = parsed.value;
	if (parsed.numFmt && format === 'General')
		cell.styleId = applyStylePatch(workbook, cell.styleId, { numFmt: parsed.numFmt });
}

/** Writes a value into a cell (no parsing) without opening an undo step. */
export function writeValue(sheet: Worksheet, row: number, col: number, value: CellValue): void {
	if (value === null || value === '') {
		const existing = getCell(sheet, row, col);
		if (existing) clearContents(existing);
		pruneCell(sheet, row, col);
		return;
	}
	const cell = ensureCell(sheet, row, col);
	clearContents(cell);
	cell.value = value;
}

export function clearContents(cell: Cell): void {
	cell.value = null;
	delete cell.formula;
	delete cell.richText;
	delete cell.arrayRange;
	delete cell.dynamicArray;
	delete cell.legacyFormula;
	// Whatever is written next is user data, not a dynamic-array result.
	delete (cell as { spillAnchor?: unknown }).spillAnchor;
}

export function setCellInput(
	ctx: EditContext,
	s: number,
	row: number,
	col: number,
	text: string,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const range = cellRange(row, col);
	ctx.run(
		`Typing in ${formatAddress({ row, col })}`,
		'cells',
		writeScopes(sheet, s, range),
		() => {
			writeInput(ctx, sheet, row, col, text);
			syncTableHeader(ctx, sheet, row, col);
		},
		{ sheet: s, ranges: [range] },
	);
}

export function setCellValue(
	ctx: EditContext,
	s: number,
	row: number,
	col: number,
	value: CellValue,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const range = cellRange(row, col);
	ctx.run(
		`Edit ${formatAddress({ row, col })}`,
		'cells',
		writeScopes(sheet, s, range),
		() => {
			writeValue(sheet, row, col, value);
			syncTableHeader(ctx, sheet, row, col);
		},
		{ sheet: s, ranges: [range] },
	);
}

export function setRangeValues(
	ctx: EditContext,
	s: number,
	at: CellAddress,
	values: CellValue[][],
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const width = Math.max(0, ...values.map((row) => row.length));
	if (!values.length || !width) return;
	const range = cellRange(at.row, at.col, at.row + values.length - 1, at.col + width - 1);
	ctx.run(
		'Edit cells',
		'cells',
		writeScopes(sheet, s, range),
		() => {
			values.forEach((rowValues, r) =>
				rowValues.forEach((value, c) => {
					writeValue(sheet, at.row + r, at.col + c, value);
					syncTableHeader(ctx, sheet, at.row + r, at.col + c);
				}),
			);
		},
		{ sheet: s, ranges: [range] },
	);
}

const CLEAR_LABELS: Record<ClearWhat, string> = {
	all: 'Clear all',
	contents: 'Clear contents',
	formats: 'Clear formats',
	comments: 'Clear comments',
	hyperlinks: 'Clear hyperlinks',
};

/** The sheet properties each kind of clear changes besides the cells. */
const CLEAR_PARTS: Record<ClearWhat, SheetPart[]> = {
	all: ['merges', 'conditionalFormats', 'comments', 'hyperlinks'],
	contents: [],
	formats: ['merges', 'conditionalFormats'],
	comments: ['comments'],
	hyperlinks: ['hyperlinks'],
};

export function clearRange(ctx: EditContext, s: number, range: CellRange, what: ClearWhat): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	const cellsChange = what === 'all' || what === 'contents' || what === 'formats';
	const scopes: EditScope[] = cellsChange
		? writeScopes(sheet, s, r, CLEAR_PARTS[what])
		: [{ kind: 'parts', sheet: s, parts: CLEAR_PARTS[what] }];
	ctx.run(
		CLEAR_LABELS[what],
		what === 'formats' ? 'format' : what === 'contents' || what === 'all' ? 'cells' : 'annotations',
		scopes,
		() => clearIn(ctx, sheet, r, what),
		{ sheet: s, ranges: [r] },
	);
}

/** Clears part of a range's content in place (no undo step). */
export function clearIn(ctx: EditContext, sheet: Worksheet, r: CellRange, what: ClearWhat): void {
	const positions: [number, number][] = [];
	forEachCellInRange(sheet, r, (_cell, row, col) => positions.push([row, col]));
	if (what === 'contents' || what === 'all' || what === 'formats') {
		for (const [row, col] of positions) {
			const cell = getCell(sheet, row, col);
			if (!cell) continue;
			if (what === 'all') deleteCell(sheet, row, col);
			else {
				if (what === 'contents') clearContents(cell);
				else delete cell.styleId;
				pruneCell(sheet, row, col);
			}
		}
		if (what === 'contents')
			for (const [row, col] of positions) syncTableHeader(ctx, sheet, row, col);
	}
	if (what === 'all' || what === 'formats') {
		sheet.merges = sheet.merges.filter((m) => !rangesIntersect(m, r));
		sheet.conditionalFormats = sheet.conditionalFormats
			.map((cf) => ({ ...cf, ranges: cf.ranges.flatMap((cr) => subtractRange(cr, r)) }))
			.filter((cf) => cf.ranges.length > 0);
	}
	if (what === 'all' || what === 'comments')
		sheet.comments = sheet.comments.filter(
			(c) => !rangesIntersect(cellRange(c.address.row, c.address.col), r),
		);
	if (what === 'all' || what === 'hyperlinks')
		sheet.hyperlinks = sheet.hyperlinks.filter((h) => !rangesIntersect(h.range, r));
}
