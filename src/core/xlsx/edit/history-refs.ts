// Undo snapshots for edits that rewrite references across the workbook (row and column shifts,
// moves). Instead of copying every cell they record the sheet metadata, the defined names and
// only the cell formulas that changed, plus the cells a shift destroyed or created.
import type { CellRange } from '../address.js';
import { getCell, putCell } from '../cells.js';
import type { Cell, DefinedName, Workbook, Worksheet } from '../model.js';
import { type AxisShift, shiftIndex } from './range-math.js';

/** Limits a shift to the cells between `lo` and `hi` on the other axis. */
export interface ShiftBand {
	lo: number;
	hi: number;
}

/** One cell's reference-bearing content; absent fields mean the cell had none. */
export interface FormulaEntry {
	sheet: number;
	row: number;
	col: number;
	formula?: string;
	arrayRange?: CellRange;
	/** The cell itself, held only between the captures before and after an edit. */
	cell?: Cell;
}

/** Everything a reference rewrite may touch, minus the cells themselves. */
export interface RefsData {
	/** Each sheet without its `rows`, cloned. */
	meta: Record<string, unknown>[];
	definedNames: DefinedName[];
	formulas: FormulaEntry[];
}

export interface ShiftSpec {
	sheet: number;
	shift: AxisShift;
	band?: ShiftBand;
}

function sheetMeta(sheet: Worksheet): Record<string, unknown> {
	const out: Record<string, unknown> = {};
	for (const [key, value] of Object.entries(sheet)) if (key !== 'rows') out[key] = value;
	return structuredClone(out);
}

/**
 * Records sheet metadata and names, and (with `withFormulas`) every formula cell, for a later
 * {@link diffRefs}. The capture after an edit skips the formulas: the diff reads them live.
 */
export function captureRefs(workbook: Workbook, withFormulas = true): RefsData {
	const formulas: FormulaEntry[] = [];
	if (withFormulas)
		// `forEach` walks the cell maps markedly faster than `for...of` on large sheets.
		workbook.sheets.forEach((sheet, s) =>
			sheet.rows.forEach((cells, row) =>
				cells.forEach((cell, col) => {
					if (cell.formula === undefined && !cell.arrayRange) return;
					const entry: FormulaEntry = { sheet: s, row, col, cell };
					if (cell.formula !== undefined) entry.formula = cell.formula;
					if (cell.arrayRange) entry.arrayRange = structuredClone(cell.arrayRange);
					formulas.push(entry);
				}),
			),
		);
	return {
		meta: workbook.sheets.map(sheetMeta),
		definedNames: structuredClone(workbook.definedNames),
		formulas,
	};
}

export function restoreRefs(workbook: Workbook, data: RefsData): void {
	data.meta.forEach((meta, s) => {
		const sheet = workbook.sheets[s];
		if (!sheet) return;
		const target = sheet as unknown as Record<string, unknown>;
		const copy = structuredClone(meta);
		for (const key of Object.keys(target)) if (key !== 'rows' && !(key in copy)) delete target[key];
		Object.assign(target, copy);
	});
	workbook.definedNames = structuredClone(data.definedNames);
	for (const entry of data.formulas) {
		const sheet = workbook.sheets[entry.sheet];
		const cell = sheet && getCell(sheet, entry.row, entry.col);
		if (!cell) continue;
		if (entry.formula === undefined) delete cell.formula;
		else cell.formula = entry.formula;
		if (entry.arrayRange) cell.arrayRange = structuredClone(entry.arrayRange);
		else delete cell.arrayRange;
	}
}

const sameRange = (a: CellRange | undefined, b: CellRange | undefined): boolean =>
	a === b ||
	(!!a &&
		!!b &&
		a.start.row === b.start.row &&
		a.start.col === b.start.col &&
		a.end.row === b.end.row &&
		a.end.col === b.end.col);

const sameEntry = (a: FormulaEntry, b: FormulaEntry): boolean =>
	a.formula === b.formula && sameRange(a.arrayRange, b.arrayRange);

/** Maps a position through a shift; undefined when the position is destroyed. */
function movePosition(
	spec: ShiftSpec | undefined,
	shift: AxisShift | undefined,
	s: number,
	row: number,
	col: number,
): [number, number] | undefined {
	if (!spec || !shift || s !== spec.sheet) return [row, col];
	const along = shift.axis === 'row' ? row : col;
	const across = shift.axis === 'row' ? col : row;
	if (spec.band && (across < spec.band.lo || across > spec.band.hi)) return [row, col];
	const moved = shiftIndex(along, shift);
	if (moved === undefined) return undefined;
	return shift.axis === 'row' ? [moved, col] : [row, moved];
}

/**
 * Trims a capture to the formulas the edit changed and returns their new state. Each recorded
 * cell is followed to where the optional shift moved it; cells the edit removed or replaced are
 * left to the `cells` snapshots (or the shift's edge cells) that cover them. Unchanged formulas
 * need no restoring, so only the changes stay in memory.
 */
export function diffRefs(workbook: Workbook, before: RefsData, spec?: ShiftSpec): FormulaEntry[] {
	const keptBefore: FormulaEntry[] = [];
	const keptAfter: FormulaEntry[] = [];
	for (const e of before.formulas) {
		const cell = e.cell;
		delete e.cell;
		const at = movePosition(spec, spec?.shift, e.sheet, e.row, e.col);
		const sheet = workbook.sheets[e.sheet];
		if (!at || !sheet || !cell || getCell(sheet, at[0], at[1]) !== cell) continue;
		const now: FormulaEntry = { sheet: e.sheet, row: at[0], col: at[1] };
		if (cell.formula !== undefined) now.formula = cell.formula;
		if (cell.arrayRange) now.arrayRange = structuredClone(cell.arrayRange);
		if (sameEntry(e, now)) continue;
		keptBefore.push(e);
		keptAfter.push(now);
	}
	before.formulas = keptBefore;
	return keptAfter;
}

/** Moves a sheet's cells along a shift without touching their content (undo and redo). */
export function moveCellsRaw(sheet: Worksheet, shift: AxisShift, band?: ShiftBand): void {
	const old = sheet.rows;
	sheet.rows = new Map();
	if (shift.axis === 'row' && !band) {
		for (const [row, cells] of old) {
			const moved = shiftIndex(row, shift);
			if (moved !== undefined) sheet.rows.set(moved, cells);
		}
		return;
	}
	for (const [row, cells] of old)
		for (const [col, cell] of cells) {
			const along = shift.axis === 'row' ? row : col;
			const across = shift.axis === 'row' ? col : row;
			const inBand = !band || (across >= band.lo && across <= band.hi);
			const moved = inBand ? shiftIndex(along, shift) : along;
			if (moved === undefined) continue;
			if (shift.axis === 'row') putCell(sheet, moved, col, cell);
			else putCell(sheet, row, moved, cell);
		}
}

/**
 * The cells a shift destroys (`phase` before: deleted or pushed off the sheet) or creates
 * (`phase` after: the inserted span), cloned.
 */
export function shiftEdgeCells(
	sheet: Worksheet,
	spec: ShiftSpec,
	phase: 'before' | 'after',
): [number, number, Cell][] {
	const { shift, band } = spec;
	const out: [number, number, Cell][] = [];
	if (phase === 'after' && shift.count < 0) return out;
	const hit = (along: number): boolean =>
		phase === 'before'
			? shiftIndex(along, shift) === undefined
			: along >= shift.at && along < shift.at + shift.count;
	for (const [row, cells] of sheet.rows) {
		if (shift.axis === 'row' && !hit(row)) continue;
		for (const [col, cell] of cells) {
			const along = shift.axis === 'row' ? row : col;
			const across = shift.axis === 'row' ? col : row;
			if (band && (across < band.lo || across > band.hi)) continue;
			if (hit(along)) out.push([row, col, structuredClone(cell)]);
		}
	}
	return out;
}
