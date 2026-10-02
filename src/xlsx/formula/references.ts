// Turning reference syntax into areas, and areas into values.
import { type CellRange, MAX_COL, MAX_ROW, normalizeRange } from '../address.js';
import type { Table } from '../model.js';
import type { FormulaAst, RefSpec, SheetPrefix, StructuredRef } from './ast.js';
import type { EvalHost, Frame } from './context.js';
import {
	type Area,
	ERR,
	fail,
	LambdaValue,
	Matrix,
	RefValue,
	type Scalar,
	type Value,
} from './values.js';

/** Sheet index for a name (case-insensitive), or -1. */
export function sheetIndex(host: EvalHost, name: string): number {
	const lower = name.toLowerCase();
	return host.workbook.sheets.findIndex((sheet) => sheet.name.toLowerCase() === lower);
}

/** The sheets a prefix names (several for a 3D reference); `undefined` when one is missing. */
export function prefixSheets(
	host: EvalHost,
	prefix: SheetPrefix | undefined,
	current: number,
): number[] | undefined {
	if (!prefix) return [current];
	if (prefix.book !== undefined) return undefined;
	const first = sheetIndex(host, prefix.sheet);
	if (first < 0) return undefined;
	if (prefix.sheet2 === undefined) return [first];
	const last = sheetIndex(host, prefix.sheet2);
	if (last < 0) return undefined;
	const out: number[] = [];
	for (let i = Math.min(first, last); i <= Math.max(first, last); i++) out.push(i);
	return out;
}

export const specRange = (spec: RefSpec): CellRange =>
	normalizeRange({
		start: { row: spec.start.row, col: spec.start.col },
		end: { row: spec.end.row, col: spec.end.col },
	});

/** Evaluates a `ref` node to a reference value (or `#REF!`). */
export function resolveRefNode(node: Extract<FormulaAst, { type: 'ref' }>, frame: Frame): Value {
	if (!node.ref) return ERR.REF;
	const sheets = prefixSheets(frame.host, node.prefix, frame.sheet);
	if (!sheets) return ERR.REF;
	const range = specRange(node.ref);
	if (node.spill) {
		const sheet = sheets[0] ?? frame.sheet;
		const spill = frame.host.spillRange(sheet, range.start.row, range.start.col);
		return spill ? new RefValue([{ sheet, range: spill }]) : ERR.REF;
	}
	return new RefValue(sheets.map((sheet) => ({ sheet, range })));
}

/** Finds a table by name in any sheet, or the table containing a cell. */
export function findTable(
	host: EvalHost,
	name: string | undefined,
	sheet: number,
	row: number,
	col: number,
): { table: Table; sheet: number } | undefined {
	const sheets = host.workbook.sheets;
	if (name) {
		const lower = name.toLowerCase();
		for (let s = 0; s < sheets.length; s++) {
			const table = sheets[s]?.tables.find(
				(t) => t.displayName.toLowerCase() === lower || t.name.toLowerCase() === lower,
			);
			if (table) return { table, sheet: s };
		}
		return undefined;
	}
	const table = sheets[sheet]?.tables.find(
		(t) =>
			row >= t.range.start.row &&
			row <= t.range.end.row &&
			col >= t.range.start.col &&
			col <= t.range.end.col,
	);
	return table ? { table, sheet } : undefined;
}

/** Resolves a structured reference to an area, or an error. */
export function resolveStructured(ref: StructuredRef, frame: Frame): Value {
	const found = findTable(frame.host, ref.table, frame.sheet, frame.row, frame.col);
	if (!found) return ref.table ? ERR.NAME : ERR.REF;
	const { table, sheet } = found;
	const range = normalizeRange(table.range);
	const headerRow = table.headerRow ? range.start.row : -1;
	const dataStart = range.start.row + (table.headerRow ? 1 : 0);
	const dataEnd = range.end.row - (table.totalsRow ? 1 : 0);
	const totalsRow = table.totalsRow ? range.end.row : -1;
	const columnOf = (name: string): number => {
		const lower = name.toLowerCase();
		const index = table.columns.findIndex((c) => c.name.toLowerCase() === lower);
		return index < 0 ? -1 : range.start.col + index;
	};
	let c1 = range.start.col;
	let c2 = range.end.col;
	if (ref.column !== undefined) {
		c1 = columnOf(ref.column);
		c2 = ref.column2 !== undefined ? columnOf(ref.column2) : c1;
		if (c1 < 0 || c2 < 0) return ERR.REF;
	}
	const specials = ref.specials.length ? ref.specials : ['#Data'];
	let r1 = Infinity;
	let r2 = -Infinity;
	const include = (a: number, b: number): void => {
		r1 = Math.min(r1, a);
		r2 = Math.max(r2, b);
	};
	for (const special of specials) {
		if (special === '#All') include(range.start.row, range.end.row);
		else if (special === '#Data') include(dataStart, dataEnd);
		else if (special === '#Headers') {
			if (headerRow < 0) return ERR.REF;
			include(headerRow, headerRow);
		} else if (special === '#Totals') {
			if (totalsRow < 0) return ERR.REF;
			include(totalsRow, totalsRow);
		} else {
			if (sheet !== frame.sheet || frame.row < dataStart || frame.row > dataEnd) return ERR.VALUE;
			include(frame.row, frame.row);
		}
	}
	if (r1 > r2) return ERR.REF;
	return new RefValue([
		{ sheet, range: normalizeRange({ start: { row: r1, col: c1 }, end: { row: r2, col: c2 } }) },
	]);
}

/** `:` between two references: the bounding box. */
export function rangeOperator(left: Value, right: Value): Value {
	if (!(left instanceof RefValue) || !(right instanceof RefValue)) return ERR.VALUE;
	const a = left.areas;
	const b = right.areas;
	if (a.length !== 1 || b.length !== 1) return ERR.VALUE;
	const x = a[0] as Area;
	const y = b[0] as Area;
	if (x.sheet !== y.sheet) return ERR.VALUE;
	return new RefValue([
		{
			sheet: x.sheet,
			range: {
				start: {
					row: Math.min(x.range.start.row, y.range.start.row),
					col: Math.min(x.range.start.col, y.range.start.col),
				},
				end: {
					row: Math.max(x.range.end.row, y.range.end.row),
					col: Math.max(x.range.end.col, y.range.end.col),
				},
			},
		},
	]);
}

/** The space operator: the cells common to both references, or `#NULL!`. */
export function intersectOperator(left: Value, right: Value): Value {
	if (!(left instanceof RefValue) || !(right instanceof RefValue)) return ERR.VALUE;
	const out: Area[] = [];
	for (const x of left.areas) {
		for (const y of right.areas) {
			if (x.sheet !== y.sheet) continue;
			const start = {
				row: Math.max(x.range.start.row, y.range.start.row),
				col: Math.max(x.range.start.col, y.range.start.col),
			};
			const end = {
				row: Math.min(x.range.end.row, y.range.end.row),
				col: Math.min(x.range.end.col, y.range.end.col),
			};
			if (start.row <= end.row && start.col <= end.col)
				out.push({ sheet: x.sheet, range: { start, end } });
		}
	}
	return out.length ? new RefValue(out) : ERR.NULL;
}

/** The comma operator inside parentheses: a multi-area reference. */
export function unionOperator(left: Value, right: Value): Value {
	if (!(left instanceof RefValue) || !(right instanceof RefValue)) return ERR.VALUE;
	return new RefValue([...left.areas, ...right.areas]);
}

/** Ranges larger than this are clipped to the sheet's used bounds when read densely. */
const CLIP_THRESHOLD = 100_000;
const MAX_MATRIX_CELLS = 4_000_000;

/** Reads an area into a dense matrix (blank cells are `null`). */
export function areaToMatrix(host: EvalHost, area: Area): Matrix {
	const { start } = area.range;
	let { row: endRow, col: endCol } = area.range.end;
	const rows = endRow - start.row + 1;
	const cols = endCol - start.col + 1;
	if (rows * cols > CLIP_THRESHOLD) {
		const bounds = host.bounds(area.sheet);
		endRow = Math.max(start.row, Math.min(endRow, bounds.rows - 1));
		endCol = Math.max(start.col, Math.min(endCol, bounds.cols - 1));
	}
	if ((endRow - start.row + 1) * (endCol - start.col + 1) > MAX_MATRIX_CELLS) fail(ERR.NUM);
	return Matrix.build(endRow - start.row + 1, endCol - start.col + 1, (r, c) =>
		host.readCell(area.sheet, start.row + r, start.col + c),
	);
}

/** A dense matrix for any value; unions and lambdas are `#VALUE!`. */
export function toMatrix(host: EvalHost, value: Value): Matrix {
	if (value instanceof Matrix) return value;
	if (value instanceof RefValue) {
		const area = value.areas[0];
		if (value.areas.length !== 1 || !area) fail(ERR.VALUE);
		return areaToMatrix(host, area);
	}
	if (value instanceof LambdaValue) fail(ERR.CALC);
	return new Matrix([[value]]);
}

/** A single value: refs by implicit intersection with the formula's cell, arrays by their first element. */
export function implicitIntersection(value: Value, frame: Frame): Scalar {
	if (value instanceof Matrix) return value.get(0, 0);
	if (value instanceof LambdaValue) return ERR.CALC;
	if (!(value instanceof RefValue)) return value;
	const area = value.areas[0];
	if (value.areas.length !== 1 || !area) return ERR.VALUE;
	const { start, end } = area.range;
	if (start.row === end.row && start.col === end.col) {
		return frame.host.readCell(area.sheet, start.row, start.col);
	}
	const rowInside = frame.row >= start.row && frame.row <= end.row;
	const colInside = frame.col >= start.col && frame.col <= end.col;
	if (start.col === end.col && rowInside)
		return frame.host.readCell(area.sheet, frame.row, start.col);
	if (start.row === end.row && colInside)
		return frame.host.readCell(area.sheet, start.row, frame.col);
	if (rowInside && colInside) return frame.host.readCell(area.sheet, frame.row, frame.col);
	return ERR.VALUE;
}

/** Whether an area covers whole columns or rows (for sparse iteration choices). */
export const isHuge = (range: CellRange): boolean =>
	range.end.row - range.start.row >= MAX_ROW || range.end.col - range.start.col >= MAX_COL;
