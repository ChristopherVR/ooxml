import type { CellRange } from '../address.js';
import type { Workbook } from '../model.js';
import * as annotations from './annotations.js';
import * as borders from './borders.js';
import * as cellShift from './cell-shift.js';
import * as values from './cell-values.js';
import * as clipboard from './clipboard.js';
import type { EditContext, RunInfo } from './context.js';
import { createCalcEngine } from './deps.js';
import * as dimensions from './dimensions.js';
import * as fill from './fill.js';
import * as filter from './filter.js';
import * as find from './find.js';
import * as format from './format.js';
import {
	type EditScope,
	History,
	type HistoryStep,
	captureScope,
	restoreSnapshot,
} from './history.js';
import { SessionCalculator } from './session-calc.js';
import * as calcMode from './calc-mode.js';
import * as drawings from './drawings.js';
import * as cellStyle from './cell-style.js';
import * as cfEdits from './conditional-formats.js';
import * as charts from './charts.js';
import * as tableEdits from './table-edits.js';
import * as duplicates from './duplicates.js';
import * as outline from './outline.js';
import * as page from './page.js';
import * as protection from './protection.js';
import * as merge from './merge.js';
import * as rowFit from './row-autofit.js';
import * as sheets from './sheets.js';
import * as sort from './sort.js';
import * as structure from './structure.js';
import * as tables from './tables.js';
import type {
	EditSession,
	EditSessionOptions,
	WorkbookChange,
	WorkbookChangeKind,
} from './types.js';
import { validateCellInput } from './validation.js';
import * as view from './view.js';

interface OpenStep {
	step: HistoryStep;
	kind: WorkbookChangeKind;
	sheet: number | undefined;
	ranges: CellRange[];
}

/** Wraps a workbook in an undoable editing session. */
export function createEditSession(
	workbook: Workbook,
	options: EditSessionOptions = {},
): EditSession {
	const calc = createCalcEngine(workbook);
	const autoRecalc = options.recalc ?? true;
	const history = new History(options.historyLimit ?? 100);
	const listeners = new Set<(change: WorkbookChange) => void>();
	let open: OpenStep | undefined;

	const emit = (change: WorkbookChange): void => {
		for (const listener of [...listeners]) listener(change);
	};

	const calculator = new SessionCalculator(workbook, calc, autoRecalc);
	const recalc = (step: HistoryStep): void => calculator.afterStep(step);

	const begin = (label: string, kind: WorkbookChangeKind, info: RunInfo): OpenStep => ({
		step: { label, entries: [], structural: info.structural ?? false },
		kind,
		sheet: info.sheet,
		ranges: [...(info.ranges ?? [])],
	});

	const finish = (current: OpenStep): void => {
		if (!current.step.entries.length) return;
		history.push(current.step);
		recalc(current.step);
		const change: WorkbookChange = {
			kind: current.kind,
			label: current.step.label,
			structural: current.step.structural,
		};
		if (current.sheet !== undefined) change.sheet = current.sheet;
		if (current.ranges.length) change.ranges = current.ranges;
		emit(change);
	};

	const rollback = (current: OpenStep, from: number): void => {
		for (let i = current.step.entries.length - 1; i >= from; i--) {
			const entry = current.step.entries[i];
			if (entry) restoreSnapshot(workbook, entry.before);
		}
		current.step.entries.length = from;
	};

	const run = <T>(
		label: string,
		kind: WorkbookChangeKind,
		scopes: EditScope[],
		fn: () => T,
		info: RunInfo = {},
	): T => {
		const outer = !open;
		const current = open ?? begin(label, kind, info);
		if (!outer) {
			if (info.structural) current.step.structural = true;
			if (current.sheet === undefined && info.sheet !== undefined) current.sheet = info.sheet;
			current.ranges.push(...(info.ranges ?? []));
		}
		open = current;
		const mark = current.step.entries.length;
		const befores = scopes.map((scope) => captureScope(workbook, scope));
		let result: T;
		try {
			result = fn();
		} catch (error) {
			for (let i = befores.length - 1; i >= 0; i--) {
				const snapshot = befores[i];
				if (snapshot) restoreSnapshot(workbook, snapshot);
			}
			current.step.entries.length = mark;
			if (outer) open = undefined;
			throw error;
		}
		scopes.forEach((scope, i) => {
			const before = befores[i];
			if (before) current.step.entries.push({ before, after: captureScope(workbook, scope) });
		});
		if (outer) {
			refitRows(current);
			open = undefined;
			finish(current);
		}
		return result;
	};

	const ctx: EditContext = { workbook, calc, run };
	if (options.defaultSheetBase) ctx.defaultSheetBase = options.defaultSheetBase;

	// Excel re-fits rows without a custom height when their text or format changes. Runs while
	// the step is still open and after its own entries, so undo restores heights first.
	const autoRows = options.autoRowHeight ?? true;
	const refitRows = (current: OpenStep): void => {
		const { sheet: s, kind, ranges } = current;
		if (!autoRows || s === undefined || !ranges.length || !current.step.entries.length) return;
		if (kind !== 'cells' && kind !== 'format' && kind !== 'batch') return;
		const sheet = workbook.sheets[s];
		if (!sheet) return;
		try {
			rowFit.autoGrowRows(ctx, s, rowFit.rowsToRefit(sheet, ranges), options.measureText);
		} catch (error) {
			const message = `Row height update failed: ${error instanceof Error ? error.message : String(error)}`;
			if (!workbook.warnings.includes(message)) workbook.warnings.push(message);
		}
	};

	const replay = (kind: 'undo' | 'redo'): boolean => {
		if (open) throw new Error(`Cannot ${kind} inside a batch`);
		const step = kind === 'undo' ? history.undo(workbook) : history.redo(workbook);
		if (!step) return false;
		recalc(step);
		emit({ kind, label: step.label, structural: step.structural });
		return true;
	};

	return {
		workbook,
		calc,
		setCellInput: (s, r, c, text) => values.setCellInput(ctx, s, r, c, text),
		setCellValue: (s, r, c, value) => values.setCellValue(ctx, s, r, c, value),
		setRangeValues: (s, at, rows) => values.setRangeValues(ctx, s, at, rows),
		clearRange: (s, range, what) => values.clearRange(ctx, s, range, what),
		applyStyle: (s, ranges, patch) => format.applyStyle(ctx, s, ranges, patch),
		setBorders: (s, range, preset, edge) => borders.setBorders(ctx, s, range, preset, edge),
		insertRows: (s, at, count) => structure.insertRows(ctx, s, at, count),
		deleteRows: (s, at, count) => structure.deleteRows(ctx, s, at, count),
		insertColumns: (s, at, count) => structure.insertColumns(ctx, s, at, count),
		deleteColumns: (s, at, count) => structure.deleteColumns(ctx, s, at, count),
		insertCellsShift: (s, range, dir) => cellShift.insertCellsShift(ctx, s, range, dir),
		deleteCellsShift: (s, range, dir) => cellShift.deleteCellsShift(ctx, s, range, dir),
		setColumnWidth: (s, cols, width, measure) =>
			dimensions.setColumnWidth(ctx, s, cols, width, measure),
		setRowHeight: (s, rows, height) => dimensions.setRowHeight(ctx, s, rows, height),
		setHidden: (s, axis, indices, hidden) => dimensions.setHidden(ctx, s, axis, indices, hidden),
		merge: (s, range, mode) => merge.merge(ctx, s, range, mode),
		unmerge: (s, range) => merge.unmerge(ctx, s, range),
		sort: (s, range, keys, hasHeader) => sort.sortRange(ctx, s, range, keys, hasHeader),
		fill: (s, source, target) => fill.fillRange(ctx, s, source, target),
		copy: (s, range) => clipboard.copyRange(workbook, s, range),
		cut: (s, range) => ({ ...clipboard.copyRange(workbook, s, range), cut: true }),
		paste: (s, at, payload, mode) => clipboard.pasteAt(ctx, s, at, payload, mode ?? 'all'),
		findAll: (query) => find.findAll(workbook, query),
		replaceAll: (query, replacement) => find.replaceAll(ctx, query, replacement),
		replaceOne: (query, replacement, at) => find.replaceOne(ctx, query, replacement, at),
		addSheet: (name, at) => sheets.addSheet(ctx, name, at),
		deleteSheet: (i) => sheets.deleteSheet(ctx, i),
		renameSheet: (i, name) => sheets.renameSheet(ctx, i, name),
		moveSheet: (from, to) => sheets.moveSheet(ctx, from, to),
		duplicateSheet: (i) => sheets.duplicateSheet(ctx, i),
		setSheetState: (i, state) => sheets.setSheetState(ctx, i, state),
		setTabColor: (i, color) => sheets.setTabColor(ctx, i, color),
		setFreeze: (s, freeze) => view.setFreeze(ctx, s, freeze),
		setComment: (s, at, text, author) => annotations.setComment(ctx, s, at, text, author),
		setHyperlink: (s, range, link) => annotations.setHyperlink(ctx, s, range, link),
		addConditionalFormat: (s, cf) => annotations.addConditionalFormat(ctx, s, cf),
		clearConditionalFormats: (s, range) => annotations.clearConditionalFormats(ctx, s, range),
		setDataValidation: (s, dv, range) => annotations.setDataValidation(ctx, s, dv, range),
		setAutoFilter: (s, range) => filter.setAutoFilter(ctx, s, range),
		filterColumn: (s, col, vals) => filter.filterColumn(ctx, s, col, vals),
		sortByColumn: (s, col, descending) => filter.sortByColumn(ctx, s, col, descending ?? false),
		createTable: (s, range, hasHeader, styleName) =>
			tables.createTable(ctx, s, range, hasHeader, styleName),
		setSheetView: (s, patch) => view.setSheetView(ctx, s, patch),
		setDefinedName: (name) => view.setDefinedName(ctx, name),
		deleteDefinedName: (name, localSheet) => view.deleteDefinedName(ctx, name, localSheet),
		setDrawingAnchor: (s, index, anchor) => drawings.setDrawingAnchor(ctx, s, index, anchor),
		deleteDrawing: (s, index) => drawings.deleteDrawing(ctx, s, index),
		addChart: (s, chart) => charts.addChart(ctx, s, chart),
		updateChart: (s, index, patch) => charts.updateChart(ctx, s, index, patch),
		addImage: (s, bytes, type, anchor, name) => charts.addImage(ctx, s, bytes, type, anchor, name),
		replaceConditionalFormat: (s, index, cf) => cfEdits.replaceConditionalFormat(ctx, s, index, cf),
		removeConditionalFormat: (s, index) => cfEdits.removeConditionalFormat(ctx, s, index),
		setConditionalRulePriority: (s, f, r, priority) =>
			cfEdits.setConditionalRulePriority(ctx, s, f, r, priority),
		moveConditionalRule: (s, f, r, direction) =>
			cfEdits.moveConditionalRule(ctx, s, f, r, direction),
		updateTable: (s, table, patch) => tableEdits.updateTable(ctx, s, table, patch),
		resizeTable: (s, table, range) => tableEdits.resizeTable(ctx, s, table, range),
		convertTableToRange: (s, table) => tableEdits.convertTableToRange(ctx, s, table),
		applyCellStyle: (s, ranges, name) => cellStyle.applyCellStyle(ctx, s, ranges, name),
		setPageSetup: (s, patch) => page.setPageSetup(ctx, s, patch),
		setPrintOptions: (s, patch) => page.setPrintOptions(ctx, s, patch),
		setSheetProtection: (s, p, password) => protection.setSheetProtection(ctx, s, p, password),
		setWorkbookProtection: (locked, password) =>
			protection.setWorkbookProtection(ctx, locked, password),
		groupRows: (s, from, to) => outline.groupRows(ctx, s, from, to),
		ungroupRows: (s, from, to) => outline.ungroupRows(ctx, s, from, to),
		groupColumns: (s, from, to) => outline.groupColumns(ctx, s, from, to),
		ungroupColumns: (s, from, to) => outline.ungroupColumns(ctx, s, from, to),
		setDefaultColumnWidth: (s, width) => outline.setDefaultColumnWidth(ctx, s, width),
		removeDuplicates: (s, range, cols, hasHeader) =>
			duplicates.removeDuplicates(ctx, s, range, cols, hasHeader),
		setCalcMode: (mode) => calcMode.setCalcMode(ctx, mode),
		setAutoRecalc: (enabled) => calcMode.setCalcMode(ctx, enabled ? 'auto' : 'manual'),
		autoRecalc: () => autoRecalc && calcMode.calcModeOf(workbook) === 'auto',
		calculateNow() {
			calculator.full();
			emit({ kind: 'cells', label: 'Calculate now', structural: false });
		},
		calculateSheet(s) {
			calculator.sheet(s);
			emit({ kind: 'cells', label: 'Calculate sheet', sheet: s, structural: false });
		},
		validate: (s, r, c, value) =>
			validateCellInput(workbook, s, r, c, value, {
				evaluate: (formula, at) => calc.evaluate(formula, at),
			}),
		batch(label, fn) {
			if (open) {
				fn();
				return;
			}
			const current = begin(label, 'batch', {});
			open = current;
			try {
				fn();
			} catch (error) {
				rollback(current, 0);
				open = undefined;
				throw error;
			}
			refitRows(current);
			open = undefined;
			finish(current);
		},
		undo: () => replay('undo'),
		redo: () => replay('redo'),
		canUndo: () => history.canUndo(),
		canRedo: () => history.canRedo(),
		undoLabel: () => history.undoLabel(),
		redoLabel: () => history.redoLabel(),
		onChange(listener) {
			listeners.add(listener);
			return () => listeners.delete(listener);
		},
	};
}
