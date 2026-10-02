import { type CellRange, normalizeRange, rangesIntersect } from '../address.js';
import { deleteCell, forEachCellInRange, putCell, usedRange } from '../cells.js';
import type { Cell, CellValue, Worksheet } from '../model.js';
import { isCellError } from '../model.js';
import { type EditContext, sheetAt } from './context.js';
import { isSpilledCell, translateFormula } from './deps.js';
import { cellRange } from './range-math.js';

export interface SortKey {
	col: number;
	descending?: boolean;
}

/** Excel's ascending sort ranks: numbers, then text, then logicals, then errors. */
function rank(value: CellValue): number {
	if (typeof value === 'number') return 0;
	if (typeof value === 'string') return 1;
	if (typeof value === 'boolean') return 2;
	if (isCellError(value)) return 3;
	return 4;
}

const collator = new Intl.Collator(undefined, { sensitivity: 'base' });

const isBlank = (value: CellValue): boolean => value === null || value === '';

/**
 * Compares two cell values as Excel's sort does. Blanks always sort last, whatever the direction;
 * text compares case-insensitively.
 */
export function compareCellValues(a: CellValue, b: CellValue, descending = false): number {
	const blankA = isBlank(a);
	const blankB = isBlank(b);
	if (blankA || blankB) return blankA === blankB ? 0 : blankA ? 1 : -1;
	const ra = rank(a);
	const rb = rank(b);
	let result: number;
	if (ra !== rb) result = ra - rb;
	else if (typeof a === 'number' && typeof b === 'number') result = a - b;
	else if (typeof a === 'string' && typeof b === 'string') result = collator.compare(a, b);
	else if (typeof a === 'boolean' && typeof b === 'boolean') result = Number(a) - Number(b);
	else result = 0;
	return descending ? -result : result;
}

/** Sorts rows of a range in place (no undo step). Exported for the auto-filter sort. */
export function sortRows(
	sheet: Worksheet,
	range: CellRange,
	keys: SortKey[],
	hasHeader: boolean,
): void {
	const r = normalizeRange(range);
	const used = usedRange(sheet);
	const lastRow = Math.min(r.end.row, used?.end.row ?? r.start.row);
	const firstRow = hasHeader ? r.start.row + 1 : r.start.row;
	if (lastRow <= firstRow) return;
	const dataArea = cellRange(firstRow, r.start.col, lastRow, r.end.col);
	if (sheet.merges.some((m) => rangesIntersect(m, dataArea)))
		throw new Error('Cannot sort a range that contains merged cells.');
	let spilled = false;
	forEachCellInRange(sheet, dataArea, (cell) => {
		spilled ||= isSpilledCell(cell);
	});
	if (spilled) throw new Error('Cannot sort a range that contains part of a dynamic array.');
	const sortKeys = keys.length ? keys : [{ col: r.start.col }];
	const rows: { row: number; cells: Map<number, Cell> }[] = [];
	for (let row = firstRow; row <= lastRow; row++) {
		const cells = new Map<number, Cell>();
		for (const [col, cell] of sheet.rows.get(row) ?? [])
			if (col >= r.start.col && col <= r.end.col) cells.set(col, cell);
		rows.push({ row, cells });
	}
	const sorted = [...rows].sort((a, b) => {
		for (const key of sortKeys) {
			const result = compareCellValues(
				a.cells.get(key.col)?.value ?? null,
				b.cells.get(key.col)?.value ?? null,
				key.descending,
			);
			if (result) return result;
		}
		return 0;
	});
	for (const { row, cells } of rows) for (const col of cells.keys()) deleteCell(sheet, row, col);
	moveAnnotations(sheet, dataArea, new Map(sorted.map((source, i) => [source.row, firstRow + i])));
	sorted.forEach((source, i) => {
		const row = firstRow + i;
		const dRow = row - source.row;
		for (const [col, cell] of source.cells) {
			if (cell.formula !== undefined && dRow) {
				try {
					cell.formula = translateFormula(cell.formula, dRow, 0);
				} catch {
					// Keep a formula the parser cannot read as it was.
				}
			}
			putCell(sheet, row, col, cell);
		}
	});
}

/**
 * Comments and single-row hyperlinks inside the sorted area travel with their rows (row
 * heights do not, as in Excel).
 */
function moveAnnotations(sheet: Worksheet, area: CellRange, rowMap: Map<number, number>): void {
	const inCols = (col: number): boolean => col >= area.start.col && col <= area.end.col;
	for (const comment of sheet.comments) {
		const next = rowMap.get(comment.address.row);
		if (next !== undefined && inCols(comment.address.col))
			comment.address = { row: next, col: comment.address.col };
	}
	for (const link of sheet.hyperlinks) {
		const { start, end } = link.range;
		const next = rowMap.get(start.row);
		if (next === undefined || start.row !== end.row || !inCols(start.col) || !inCols(end.col))
			continue;
		link.range = { start: { row: next, col: start.col }, end: { row: next, col: end.col } };
	}
}

/** Sorts the rows of `range` by `keys` (absolute column indices), stable, header kept in place. */
export function sortRange(
	ctx: EditContext,
	s: number,
	range: CellRange,
	keys: SortKey[],
	hasHeader: boolean,
): void {
	const sheet = sheetAt(ctx.workbook, s);
	const r = normalizeRange(range);
	ctx.run(
		'Sort',
		'cells',
		[{ kind: 'sheet', sheet: s }],
		() => sortRows(sheet, r, keys, hasHeader),
		{ sheet: s, ranges: [r] },
	);
}
