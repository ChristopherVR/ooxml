import { type CellRange, cellKey, normalizeRange } from '../address';
import { deleteCell, forEachCellInRange, getCell, putCell } from '../cells';
import type { Cell, DefinedName, Workbook, Worksheet } from '../model';
import {
	type RefsData,
	type ShiftSpec,
	captureRefs,
	diffRefs,
	moveCellsRaw,
	restoreRefs,
	shiftEdgeCells,
} from './history-refs';
import { cloneCell } from './clone-cell';
import { type RowInfoEntries, captureRowInfo, restoreRowInfo } from './history-rows';

/** Worksheet properties a `parts` scope can record (everything but the cells). */
export type SheetPart = Exclude<keyof Worksheet, 'rows'>;

/**
 * What an edit may change. `cells` records only the stored cells inside the ranges, `parts` some
 * of a sheet's non-cell properties (tables, merges, ...), `rowInfo` the row metadata (height,
 * hiding, outline) of the given rows only, `sheet` the whole worksheet, `refs` what
 * a workbook-wide reference rewrite touches (sheet metadata, names and the formulas that change),
 * `shift` a row or column insert or delete on one sheet (the moved cells are not copied, only the
 * destroyed and created ones and the changed references), `workbook` every sheet plus the
 * workbook-level lists (names, active sheet).
 */
export type EditScope =
	| { kind: 'cells'; sheet: number; ranges: CellRange[] }
	| { kind: 'parts'; sheet: number; parts: SheetPart[] }
	| { kind: 'rowInfo'; sheet: number; rows: readonly number[] }
	| { kind: 'sheet'; sheet: number }
	| { kind: 'refs' }
	| ({ kind: 'shift' } & ShiftSpec)
	| { kind: 'meta' }
	| { kind: 'workbook' };

type Snapshot =
	| { kind: 'cells'; sheet: number; ranges: CellRange[]; cells: [number, number, Cell][] }
	| { kind: 'parts'; sheet: number; parts: SheetPart[]; data: Partial<Worksheet> }
	| { kind: 'rowInfo'; sheet: number; rows: RowInfoEntries }
	| { kind: 'sheet'; sheet: number; data: Worksheet }
	| { kind: 'refs'; data: RefsData; activeSheet: number }
	| {
			kind: 'shift';
			spec: ShiftSpec;
			phase: 'before' | 'after';
			cells: [number, number, Cell][];
			data: RefsData;
			activeSheet: number;
	  }
	| { kind: 'meta'; definedNames: DefinedName[]; activeSheet: number; settings?: BookSettings }
	| {
			kind: 'workbook';
			sheets: Worksheet[];
			definedNames: DefinedName[];
			activeSheet: number;
			settings?: BookSettings;
	  };

/** Workbook-level settings `meta` and `workbook` snapshots also record. */
type BookSettings = Pick<
	Workbook,
	| 'structureLocked'
	| 'workbookPasswordHash'
	| 'workbookModernHash'
	| 'calcMode'
	| 'properties'
	| 'slicers'
	| 'customViews'
>;
const SETTING_KEYS = [
	'structureLocked',
	'workbookPasswordHash',
	'workbookModernHash',
	'calcMode',
	// Replaced, never mutated, by `setDocumentProperties`, so the reference is the snapshot.
	'properties',
	// Replaced, never mutated, by the slicer edits.
	'slicers',
	// Replaced, never mutated, by the custom view edits.
	'customViews',
] as const;

function captureSettings(workbook: Workbook): BookSettings {
	const out: Record<string, unknown> = {};
	for (const key of SETTING_KEYS) if (workbook[key] !== undefined) out[key] = workbook[key];
	return out as BookSettings;
}

function restoreSettings(workbook: Workbook, settings: BookSettings | undefined): void {
	if (!settings) return;
	const target = workbook as unknown as Record<string, unknown>;
	for (const key of SETTING_KEYS) {
		if (settings[key] === undefined) delete target[key];
		else target[key] = settings[key];
	}
}

export interface HistoryEntry {
	before: Snapshot;
	after: Snapshot;
}

/**
 * How the calc engine can follow a structural step without rebuilding its graph: rows or columns
 * inserted or deleted on a sheet, or a sheet renamed. Only set on a step made by one command.
 */
export type CalcHint =
	| { kind: 'shift'; sheet: number; shift: { axis: 'row' | 'col'; at: number; count: number } }
	| { kind: 'rename'; from: string; to: string };

export interface HistoryStep {
	label: string;
	entries: HistoryEntry[];
	structural: boolean;
	calc?: CalcHint;
}

/** Positions a small range covers are probed directly; larger ones scan the stored rows. */
const PROBE_LIMIT = 4096;

function captureCells(sheet: Worksheet, ranges: CellRange[]): [number, number, Cell][] {
	const cells: [number, number, Cell][] = [];
	// One range visits each cell once; only overlapping ranges can repeat one.
	const seen = ranges.length > 1 ? new Set<number>() : undefined;
	const take = (cell: Cell, row: number, col: number): void => {
		if (seen) {
			const key = cellKey(row, col);
			if (seen.has(key)) return;
			seen.add(key);
		}
		cells.push([row, col, cloneCell(cell)]);
	};
	for (const raw of ranges) {
		const range = normalizeRange(raw);
		const area = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
		if (area > PROBE_LIMIT || area > sheet.rows.size) forEachCellInRange(sheet, range, take);
		else
			for (let row = range.start.row; row <= range.end.row; row++)
				for (let col = range.start.col; col <= range.end.col; col++) {
					const cell = getCell(sheet, row, col);
					if (cell) take(cell, row, col);
				}
	}
	return cells;
}

function captureShift(
	workbook: Workbook,
	scope: { kind: 'shift' } & ShiftSpec,
	before: Snapshot | undefined,
): Snapshot {
	const spec: ShiftSpec = { sheet: scope.sheet, shift: scope.shift };
	if (scope.band) spec.band = scope.band;
	const phase = before ? 'after' : 'before';
	const sheet = workbook.sheets[scope.sheet];
	if (!sheet) throw new RangeError(`No sheet at index ${scope.sheet}`);
	const data = captureRefs(workbook, !before);
	if (before?.kind === 'shift') data.formulas = diffRefs(workbook, before.data, spec);
	const cells = shiftEdgeCells(sheet, spec, phase);
	return { kind: 'shift', spec, phase, cells, data, activeSheet: workbook.activeSheet };
}

/**
 * Records a scope. For `refs` and `shift` scopes pass the `before` snapshot when capturing the
 * state after the edit: both are then trimmed to the references that actually changed.
 */
export function captureScope(workbook: Workbook, scope: EditScope, before?: Snapshot): Snapshot {
	if (scope.kind === 'refs') {
		const data = captureRefs(workbook, !before);
		if (before?.kind === 'refs') data.formulas = diffRefs(workbook, before.data);
		return { kind: 'refs', data, activeSheet: workbook.activeSheet };
	}
	if (scope.kind === 'shift') return captureShift(workbook, scope, before);
	if (scope.kind === 'meta')
		return {
			kind: 'meta',
			definedNames: structuredClone(workbook.definedNames),
			activeSheet: workbook.activeSheet,
			settings: captureSettings(workbook),
		};
	if (scope.kind === 'workbook')
		return {
			kind: 'workbook',
			sheets: structuredClone(workbook.sheets),
			definedNames: structuredClone(workbook.definedNames),
			activeSheet: workbook.activeSheet,
			settings: captureSettings(workbook),
		};
	const sheet = workbook.sheets[scope.sheet];
	if (!sheet) throw new RangeError(`No sheet at index ${scope.sheet}`);
	if (scope.kind === 'sheet')
		return { kind: 'sheet', sheet: scope.sheet, data: structuredClone(sheet) };
	if (scope.kind === 'rowInfo')
		return { kind: 'rowInfo', sheet: scope.sheet, rows: captureRowInfo(sheet, scope.rows) };
	if (scope.kind === 'parts') {
		const data: Record<string, unknown> = {};
		for (const part of scope.parts) if (sheet[part] !== undefined) data[part] = sheet[part];
		return {
			kind: 'parts',
			sheet: scope.sheet,
			parts: [...scope.parts],
			data: structuredClone(data) as Partial<Worksheet>,
		};
	}
	return {
		kind: 'cells',
		sheet: scope.sheet,
		ranges: scope.ranges,
		cells: captureCells(sheet, scope.ranges),
	};
}

/** Undoes or redoes a row or column shift recorded by a `shift` snapshot. */
function restoreShift(workbook: Workbook, snapshot: Extract<Snapshot, { kind: 'shift' }>): void {
	const { spec } = snapshot;
	const sheet = workbook.sheets[spec.sheet];
	if (!sheet) return;
	const shift =
		snapshot.phase === 'before' ? { ...spec.shift, count: -spec.shift.count } : spec.shift;
	moveCellsRaw(sheet, shift, spec.band);
	for (const [row, col, cell] of snapshot.cells) putCell(sheet, row, col, cloneCell(cell));
	restoreRefs(workbook, snapshot.data);
	workbook.activeSheet = snapshot.activeSheet;
}

/** Puts a snapshot back. Sheets are restored in place so references held by views stay valid. */
export function restoreSnapshot(workbook: Workbook, snapshot: Snapshot): void {
	if (snapshot.kind === 'shift') return restoreShift(workbook, snapshot);
	if (snapshot.kind === 'refs') {
		restoreRefs(workbook, snapshot.data);
		workbook.activeSheet = snapshot.activeSheet;
		return;
	}
	if (snapshot.kind === 'meta') {
		workbook.definedNames = structuredClone(snapshot.definedNames);
		workbook.activeSheet = snapshot.activeSheet;
		restoreSettings(workbook, snapshot.settings);
		return;
	}
	if (snapshot.kind === 'workbook') {
		const sheets = structuredClone(snapshot.sheets);
		workbook.sheets.length = 0;
		workbook.sheets.push(...sheets);
		workbook.definedNames = structuredClone(snapshot.definedNames);
		workbook.activeSheet = snapshot.activeSheet;
		restoreSettings(workbook, snapshot.settings);
		return;
	}
	const sheet = workbook.sheets[snapshot.sheet];
	if (!sheet) return;
	if (snapshot.kind === 'rowInfo') return restoreRowInfo(sheet, snapshot.rows);
	if (snapshot.kind === 'parts') {
		const target = sheet as unknown as Record<string, unknown>;
		const data = structuredClone(snapshot.data) as Record<string, unknown>;
		for (const part of snapshot.parts)
			if (part in data) target[part] = data[part];
			else delete target[part];
		return;
	}
	if (snapshot.kind === 'sheet') {
		const data = structuredClone(snapshot.data) as unknown as Record<string, unknown>;
		const target = sheet as unknown as Record<string, unknown>;
		for (const key of Object.keys(target)) if (!(key in data)) delete target[key];
		Object.assign(target, data);
		return;
	}
	for (const range of snapshot.ranges) {
		const doomed: [number, number][] = [];
		forEachCellInRange(sheet, range, (_cell, row, col) => doomed.push([row, col]));
		for (const [row, col] of doomed) deleteCell(sheet, row, col);
	}
	for (const [row, col, cell] of snapshot.cells) putCell(sheet, row, col, cloneCell(cell));
}

type Position = { sheet: number; row: number; col: number };

/** Whether captured cells are in row-major order, as one capture of one range leaves them. */
function rowMajor(cells: readonly [number, number, Cell][]): boolean {
	for (let i = 1; i < cells.length; i++) {
		const [row0, col0] = cells[i - 1]!;
		const [row1, col1] = cells[i]!;
		if (row1 < row0 || (row1 === row0 && col1 <= col0)) return false;
	}
	return true;
}

/**
 * The cells a step's snapshots cover, for incremental recalculation: each once, in snapshot order.
 * A bulk write is one range on one sheet, captured before and after in row-major order, so its two
 * lists are merged in step without a set of every cell.
 */
export function touchedCells(step: HistoryStep): Position[] {
	const out: Position[] = [];
	const lists = step.entries
		.flatMap((entry) => [entry.before, entry.after])
		.filter((snapshot) => snapshot.kind === 'cells' && snapshot.cells.length > 0);
	const [first, second, ...rest] = lists as Extract<Snapshot, { kind: 'cells' }>[];
	if (first && !rest.length && (!second || second.sheet === first.sheet)) {
		const sheet = first.sheet;
		for (const [row, col] of first.cells) out.push({ sheet, row, col });
		if (!second) return out;
		if (rowMajor(first.cells) && rowMajor(second.cells)) {
			let i = 0;
			for (const [row, col] of second.cells) {
				const key = cellKey(row, col);
				while (i < first.cells.length) {
					const [r, c] = first.cells[i]!;
					if (cellKey(r, c) >= key) break;
					i++;
				}
				const at = first.cells[i];
				if (at && cellKey(at[0], at[1]) === key) continue;
				out.push({ sheet, row, col });
			}
			return out;
		}
		out.length = 0;
	}
	const seen = new Map<number, Set<number>>();
	for (const entry of step.entries)
		for (const snapshot of [entry.before, entry.after]) {
			if (snapshot.kind !== 'cells') continue;
			let keys = seen.get(snapshot.sheet);
			if (!keys) seen.set(snapshot.sheet, (keys = new Set()));
			for (const [row, col] of snapshot.cells) {
				const key = cellKey(row, col);
				if (keys.has(key)) continue;
				keys.add(key);
				out.push({ sheet: snapshot.sheet, row, col });
			}
		}
	return out;
}

/** Undo and redo stacks of snapshot steps. */
export class History {
	private readonly done: HistoryStep[] = [];
	private readonly undone: HistoryStep[] = [];

	constructor(private readonly limit: number) {}

	push(step: HistoryStep): void {
		this.done.push(step);
		if (this.done.length > this.limit) this.done.shift();
		this.undone.length = 0;
	}

	undo(workbook: Workbook): HistoryStep | undefined {
		const step = this.done.pop();
		if (!step) return undefined;
		for (let i = step.entries.length - 1; i >= 0; i--) {
			const entry = step.entries[i];
			if (entry) restoreSnapshot(workbook, entry.before);
		}
		this.undone.push(step);
		return step;
	}

	redo(workbook: Workbook): HistoryStep | undefined {
		const step = this.undone.pop();
		if (!step) return undefined;
		for (const entry of step.entries) restoreSnapshot(workbook, entry.after);
		this.done.push(step);
		return step;
	}

	canUndo = (): boolean => this.done.length > 0;
	canRedo = (): boolean => this.undone.length > 0;
	undoLabel = (): string | undefined => this.done.at(-1)?.label;
	redoLabel = (): string | undefined => this.undone.at(-1)?.label;
}
