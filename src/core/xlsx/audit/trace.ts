// Trace Precedents / Trace Dependents: the arrows Excel draws for the direct references of a
// formula, one more level per repeated command. Reads the formula text with the shared tokenizer.
import { type CellRange, MAX_COL, MAX_ROW, normalizeRange, rangeContains } from '../address';
import { forEachCell, forEachCellInRange, usedRange } from '../cells';
import { safeTokens } from '../formula/ref-text';
import { isCellError, type Workbook } from '../model';

export interface CellRef {
	sheet: number;
	row: number;
	col: number;
}

export type TraceKind = 'precedents' | 'dependents';

export interface TraceArrow {
	kind: TraceKind;
	/** The cell the user traced from; arrows of one trace share it. */
	root: CellRef;
	/** 1 for direct references, one more for each repeated command. */
	level: number;
	/** The cell or range the arrow starts at. */
	from: { sheet: number; range: CellRange };
	/** The cell the arrow points at. */
	to: CellRef;
	/** The other end is on another sheet (drawn dashed with a sheet icon). */
	offSheet: boolean;
	/** The error source: drawn red. */
	error: boolean;
}

export type TraceOutcome = 'ok' | 'no-formula' | 'no-more' | 'no-dependents';

const sameCell = (a: CellRef, b: CellRef): boolean =>
	a.sheet === b.sheet && a.row === b.row && a.col === b.col;

const rangeKey = (r: CellRange): string =>
	`${r.start.row},${r.start.col},${r.end.row},${r.end.col}`;
const rootKey = (c: CellRef): string => `${c.sheet}:${c.row}:${c.col}`;
const arrowKey = (a: TraceArrow): string =>
	`${a.from.sheet}:${rangeKey(a.from.range)}>${a.to.sheet}:${a.to.row},${a.to.col}`;

function sheetIndexOf(workbook: Workbook, name: string): number {
	const lower = name.toLowerCase();
	return workbook.sheets.findIndex((s) => s.name.toLowerCase() === lower);
}

function clampToUsed(workbook: Workbook, sheet: number, range: CellRange): CellRange {
	const ws = workbook.sheets[sheet];
	const used = ws ? usedRange(ws) : undefined;
	if (!used) return { start: range.start, end: range.start };
	return {
		start: {
			row: Math.max(range.start.row, used.start.row),
			col: Math.max(range.start.col, used.start.col),
		},
		end: {
			row: Math.min(range.end.row, used.end.row),
			col: Math.min(range.end.col, used.end.col),
		},
	};
}

/** The ranges a formula reads directly (A1 references; names, structured and 3D ones are skipped). */
export function formulaReferences(
	workbook: Workbook,
	formula: string,
	sheet: number,
): { sheet: number; range: CellRange }[] {
	const tokens = safeTokens(formula);
	if (!tokens) return [];
	const own = workbook.sheets[sheet]?.name ?? '';
	const out: { sheet: number; range: CellRange }[] = [];
	for (const token of tokens) {
		if (token.kind !== 'ref' || !token.ref) continue;
		const prefix = token.prefix;
		if (prefix && (prefix.book !== undefined || prefix.sheet2 !== undefined)) continue;
		const index = sheetIndexOf(workbook, prefix?.sheet ?? own);
		if (index < 0) continue;
		const { start, end, kind } = token.ref;
		const range = normalizeRange({
			start: { row: kind === 'cols' ? 0 : start.row, col: kind === 'rows' ? 0 : start.col },
			end: {
				row: kind === 'cols' ? MAX_ROW : end.row,
				col: kind === 'rows' ? MAX_COL : end.col,
			},
		});
		const whole = kind === 'cols' || kind === 'rows';
		out.push({ sheet: index, range: whole ? clampToUsed(workbook, index, range) : range });
	}
	return out;
}

function rangeHasError(workbook: Workbook, sheet: number, range: CellRange): boolean {
	const ws = workbook.sheets[sheet];
	if (!ws) return false;
	let found = false;
	forEachCellInRange(ws, range, (cell) => {
		if (isCellError(cell.value)) found = true;
	});
	return found;
}

/** The distinct areas a formula cell reads, in formula order. */
export function directPrecedents(
	workbook: Workbook,
	cell: CellRef,
): { sheet: number; range: CellRange }[] {
	const formula = workbook.sheets[cell.sheet]?.rows.get(cell.row)?.get(cell.col)?.formula;
	if (!formula) return [];
	const seen = new Set<string>();
	return formulaReferences(workbook, formula, cell.sheet).filter((r) => {
		const key = `${r.sheet}:${rangeKey(r.range)}`;
		if (seen.has(key)) return false;
		seen.add(key);
		return true;
	});
}

/** The formula cells (on any sheet) that read `cell` directly, in sheet then row-major order. */
export function directDependents(workbook: Workbook, cell: CellRef): CellRef[] {
	const out: CellRef[] = [];
	workbook.sheets.forEach((ws, sheet) => {
		forEachCell(ws, (c, row, col) => {
			if (!c.formula) return;
			if (sheet === cell.sheet && row === cell.row && col === cell.col) return;
			const hit = formulaReferences(workbook, c.formula, sheet).some(
				(r) => r.sheet === cell.sheet && rangeContains(r.range, cell),
			);
			if (hit) out.push({ sheet, row, col });
		});
	});
	return out;
}

const single = (c: CellRef): CellRange => ({
	start: { row: c.row, col: c.col },
	end: { row: c.row, col: c.col },
});

function formulaCellsIn(workbook: Workbook, sheet: number, range: CellRange): CellRef[] {
	const ws = workbook.sheets[sheet];
	const out: CellRef[] = [];
	if (ws)
		forEachCellInRange(ws, range, (c, row, col) => {
			if (c.formula) out.push({ sheet, row, col });
		});
	return out;
}

/**
 * Adds the next level of arrows for `kind` from `anchor`: the direct references the first time,
 * the references of those cells on a repeat. Returns the extended list (unchanged when there is
 * nothing to add) and why nothing was added.
 */
export function traceNext(
	workbook: Workbook,
	arrows: readonly TraceArrow[],
	kind: TraceKind,
	anchor: CellRef,
): { arrows: TraceArrow[]; outcome: TraceOutcome } {
	const mine = arrows.filter((a) => a.kind === kind && sameCell(a.root, anchor));
	const level = mine.reduce((max, a) => Math.max(max, a.level), 0);
	const top = mine.filter((a) => a.level === level && !a.offSheet);
	let frontier: CellRef[];
	if (level === 0) frontier = [anchor];
	else if (kind === 'precedents')
		frontier = top.flatMap((a) => formulaCellsIn(workbook, a.from.sheet, a.from.range));
	else frontier = top.map((a) => a.to);
	const added: TraceArrow[] = [];
	const known = new Set(mine.map(arrowKey));
	const push = (arrow: TraceArrow): void => {
		const key = arrowKey(arrow);
		if (known.has(key)) return;
		known.add(key);
		added.push(arrow);
	};
	for (const cell of frontier) {
		if (kind === 'precedents') {
			for (const p of directPrecedents(workbook, cell))
				push({
					kind,
					root: anchor,
					level: level + 1,
					from: p,
					to: cell,
					offSheet: p.sheet !== cell.sheet,
					error: rangeHasError(workbook, p.sheet, p.range),
				});
		} else {
			for (const d of directDependents(workbook, cell)) {
				const value = workbook.sheets[d.sheet]?.rows.get(d.row)?.get(d.col)?.value;
				push({
					kind,
					root: anchor,
					level: level + 1,
					from: { sheet: cell.sheet, range: single(cell) },
					to: d,
					offSheet: d.sheet !== cell.sheet,
					error: isCellError(value),
				});
			}
		}
	}
	if (added.length) return { arrows: [...arrows, ...added], outcome: 'ok' };
	if (level > 0) return { arrows: [...arrows], outcome: 'no-more' };
	return { arrows: [...arrows], outcome: kind === 'precedents' ? 'no-formula' : 'no-dependents' };
}

/**
 * Remove Arrows: every arrow, or the outermost level of the traces of one kind (those rooted at
 * `anchor`, or all of that kind when `anchor` has none), as Excel removes them one level at a time.
 */
export function removeArrows(
	arrows: readonly TraceArrow[],
	mode: 'all' | TraceKind,
	anchor?: CellRef,
): TraceArrow[] {
	if (mode === 'all') return [];
	const ofKind = arrows.filter((a) => a.kind === mode);
	const rooted = anchor ? ofKind.filter((a) => sameCell(a.root, anchor)) : [];
	const scope = rooted.length ? rooted : ofKind;
	const tops = new Map<string, number>();
	for (const a of scope)
		tops.set(rootKey(a.root), Math.max(tops.get(rootKey(a.root)) ?? 0, a.level));
	return arrows.filter((a) => {
		if (a.kind !== mode) return true;
		const top = tops.get(rootKey(a.root));
		return top === undefined || a.level !== top;
	});
}
