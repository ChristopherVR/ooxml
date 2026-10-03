import type { CellRange } from '../address.js';
import { deleteCell, forEachCellInRange, putCell } from '../cells.js';
import type { Cell, DefinedName, Workbook, Worksheet } from '../model.js';

/**
 * What an edit may change. `cells` records only the stored cells inside the ranges, `sheet` the
 * whole worksheet, `workbook` every sheet plus the workbook-level lists (names, active sheet).
 */
export type EditScope =
	| { kind: 'cells'; sheet: number; ranges: CellRange[] }
	| { kind: 'sheet'; sheet: number }
	| { kind: 'meta' }
	| { kind: 'workbook' };

type Snapshot =
	| { kind: 'cells'; sheet: number; ranges: CellRange[]; cells: [number, number, Cell][] }
	| { kind: 'sheet'; sheet: number; data: Worksheet }
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
	'structureLocked' | 'workbookPasswordHash' | 'workbookModernHash' | 'calcMode'
>;
const SETTING_KEYS = [
	'structureLocked',
	'workbookPasswordHash',
	'workbookModernHash',
	'calcMode',
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

export interface HistoryStep {
	label: string;
	entries: HistoryEntry[];
	structural: boolean;
}

export function captureScope(workbook: Workbook, scope: EditScope): Snapshot {
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
	const cells: [number, number, Cell][] = [];
	const seen = new Set<string>();
	for (const range of scope.ranges)
		forEachCellInRange(sheet, range, (cell, row, col) => {
			const key = `${row},${col}`;
			if (seen.has(key)) return;
			seen.add(key);
			cells.push([row, col, structuredClone(cell)]);
		});
	return { kind: 'cells', sheet: scope.sheet, ranges: scope.ranges, cells };
}

/** Puts a snapshot back. Sheets are restored in place so references held by views stay valid. */
export function restoreSnapshot(workbook: Workbook, snapshot: Snapshot): void {
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
	for (const [row, col, cell] of snapshot.cells) putCell(sheet, row, col, structuredClone(cell));
}

/** The cells a step's snapshots cover, for incremental recalculation. */
export function touchedCells(step: HistoryStep): { sheet: number; row: number; col: number }[] {
	const out: { sheet: number; row: number; col: number }[] = [];
	const seen = new Set<string>();
	for (const entry of step.entries)
		for (const snapshot of [entry.before, entry.after]) {
			if (snapshot.kind !== 'cells') continue;
			for (const [row, col] of snapshot.cells) {
				const key = `${snapshot.sheet},${row},${col}`;
				if (seen.has(key)) continue;
				seen.add(key);
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
