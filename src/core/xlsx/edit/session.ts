import type { CellRange } from '../address';
import { prepareConditionalFormatFormulas } from '../formula/prepare-filterxml';
import type { ConditionalFormat, Workbook } from '../model';
import * as annotations from './annotations';
import * as borders from './borders';
import * as cellShift from './cell-shift';
import * as values from './cell-values';
import * as pictures from './cell-pictures';
import * as clipboard from './clipboard';
import { copyAreas } from './clipboard-areas';
import { pasteSelection } from './paste-selection';
import type { EditContext, RunInfo } from './context';
import { createCalcEngine, type CellPosition } from './deps';
import { applyExternal } from './external';
import * as dimensions from './dimensions';
import * as fill from './fill';
import * as filter from './filter';
import * as find from './find';
import * as format from './format';
import {
	type EditScope,
	History,
	type HistoryStep,
	captureScope,
	restoreSnapshot,
} from './history';
import { SessionCalculator } from './session-calc';
import * as calcMode from './calc-mode';
import * as dataTable from './data-table';
import * as goalSeek from './goal-seek';
import * as scenarios from './scenarios';
import { mergeScenarios } from './scenario-merge';
import * as customViews from './custom-views';
import * as sparklineGrouping from './sparkline-grouping';
import * as docProps from './doc-properties';
import * as themeEdits from './theme';
import { setSheetBackground } from './background';
import * as nativeShapes from './native-shapes';
import * as groups from './group-drawings';
import * as arrange from './arrange-drawings';
import * as drawings from './drawings';
import { insertLocalObject } from './local-object';
import * as sparklines from './sparklines';
import * as pivots from './pivots';
import { refreshAllPivotTables, refreshPivotTable } from './pivot-refresh';
import * as pivotCharts from './pivot-charts';
import * as slicers from './slicers';
import * as timelines from './timelines';
import { rangesIntersect } from '../address';
import * as cellStyle from './cell-style';
import * as cfEdits from './conditional-formats';
import * as charts from './charts';
import * as chartEx from './chart-ex';
import * as smartArt from './smart-art';
import * as tableEdits from './table-edits';
import * as duplicates from './duplicates';
import * as subtotal from './subtotal';
import * as consolidate from './consolidate';
import * as flashFillOp from './flash-fill';
import * as advancedFilter from './advanced-filter';
import * as outline from './outline';
import * as page from './page';
import * as pageBreaks from './page-breaks';
import * as printTitles from './print-titles';
import * as protection from './protection';
import * as merge from './merge';
import * as rowFit from './row-autofit';
import * as sheets from './sheets';
import * as sort from './sort';
import * as structure from './structure';
import * as tables from './tables';
import type { EditSession, EditSessionOptions, WorkbookChange, WorkbookChangeKind } from './types';
import { validateCellInput } from './validation';
import * as view from './view';

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
	const waitingForFunctions = new Map<Promise<void>, boolean>();
	const waitForFunction = (ready: Promise<void>, recalculate: boolean): void => {
		const alreadyWaiting = waitingForFunctions.has(ready);
		if (recalculate || !alreadyWaiting) waitingForFunctions.set(ready, recalculate);
		if (alreadyWaiting) return;
		void ready.then(() => {
			const recalculate = waitingForFunctions.get(ready);
			waitingForFunctions.delete(ready);
			if (recalculate) {
				calculator.finishDeferredCalculation();
				refreshTables();
			}
			emit({
				kind: recalculate ? 'cells' : 'view',
				label: 'Calculate now',
				structural: false,
				...(recalculate ? { deferredCalculation: true } : { viewOnly: true }),
			});
		});
	};
	const calc = createCalcEngine(workbook, {
		...options.calcOptions,
		waitForFunction: (ready) => waitForFunction(ready, true),
		waitForAdHocFunction: (ready) => waitForFunction(ready, false),
	});
	const autoRecalc = options.recalc ?? true;
	const history = new History(options.historyLimit ?? 100);
	const listeners = new Set<(change: WorkbookChange) => void>();
	let open: OpenStep | undefined;

	const emit = (change: WorkbookChange): void => {
		for (const listener of [...listeners]) listener(change);
	};

	const calculator = new SessionCalculator(workbook, calc, autoRecalc);
	const recalc = (step: HistoryStep, mode: 'apply' | 'undo' | 'redo' = 'apply'): void => {
		calculator.afterStep(step, mode);
		refreshTables();
	};
	// Data table results follow their inputs after every calculation (not in manual mode).
	const refreshTables = (): void => {
		if (!autoRecalc || calcMode.calcModeOf(workbook) === 'manual') return;
		if (!workbook.sheets.some((sheet) => sheet.dataTables?.length)) return;
		try {
			dataTable.refreshDataTables(workbook, calc);
		} catch (error) {
			const message = `Data table update failed: ${error instanceof Error ? error.message : String(error)}`;
			if (!workbook.warnings.includes(message)) workbook.warnings.push(message);
		}
	};

	const begin = (label: string, kind: WorkbookChangeKind, info: RunInfo): OpenStep => {
		const step: HistoryStep = { label, entries: [], structural: info.structural ?? false };
		if (info.calc) step.calc = info.calc;
		return { step, kind, sheet: info.sheet, ranges: [...(info.ranges ?? [])] };
	};

	const finish = (current: OpenStep): void => {
		if (!current.step.entries.length) return;
		history.push(current.step);
		recalc(current.step);
		const change: WorkbookChange = {
			kind: current.kind,
			label: current.step.label,
			structural: !!current.step.structural,
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
		fn: (formulaCells?: readonly CellPosition[]) => T,
		info: RunInfo = {},
	): T => {
		if (kind === 'cells')
			for (const scope of scopes)
				if (
					scope.kind === 'cells' &&
					workbook.sheets[scope.sheet]?.pivotTables?.some((pivot) =>
						scope.ranges.some((range) => rangesIntersect(range, pivot.location)),
					)
				)
					throw new Error(
						'A PivotTable result cannot be edited directly. Refresh or remove the PivotTable instead.',
					);
		for (const scope of scopes)
			if (
				kind === 'cells' &&
				scope.kind === 'cells' &&
				scope.ranges.some((range) => {
					const sheet = workbook.sheets[scope.sheet];
					return !!sheet?.dataTables?.length && dataTable.cutsDataTable(sheet, range);
				})
			)
				throw new Error(dataTable.DATA_TABLE_PART_ERROR);
		const outer = !open;
		const current = open ?? begin(label, kind, info);
		if (!outer) {
			// A hint describes one command; a step made of several cannot be followed in place.
			delete current.step.calc;
			if (info.structural && current.step.structural !== true)
				current.step.structural = info.structural;
			if (current.sheet === undefined && info.sheet !== undefined) current.sheet = info.sheet;
			current.ranges.push(...(info.ranges ?? []));
		}
		// A structural edit the engine can follow needs the graph of the workbook before it.
		if (outer && info.calc) calculator.beforeStructural();
		open = current;
		const mark = current.step.entries.length;
		const befores = scopes.map((scope) => captureScope(workbook, scope));
		let result: T;
		try {
			result = fn(befores.find((snapshot) => snapshot.kind === 'refs')?.data.formulas);
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
			if (before)
				current.step.entries.push({ before, after: captureScope(workbook, scope, before) });
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
	const prepareConditionalFormat = (format: ConditionalFormat): void => {
		const ready = prepareConditionalFormatFormulas(workbook, format);
		if (ready) waitForFunction(ready, false);
	};

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
		recalc(step, kind);
		emit({ kind, label: step.label, structural: !!step.structural });
		return true;
	};

	const session: EditSession = {
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
		sort: (s, range, keys, hasHeader, options) =>
			sort.sortRange(ctx, s, range, keys, hasHeader, options),
		fill: (s, source, target, mode) => fill.fillRange(ctx, s, source, target, mode),
		flashFill: (s, target, options) => flashFillOp.flashFill(ctx, s, target, options),
		copy: (s, range) => clipboard.copyRange(workbook, s, range),
		copyAreas: (s, ranges) => copyAreas(workbook, s, ranges),
		cut: (s, range) => ({ ...clipboard.copyRange(workbook, s, range), cut: true }),
		paste: (s, at, payload, mode) =>
			'start' in at
				? pasteSelection(ctx, s, at, payload, mode ?? 'all')
				: clipboard.pasteAt(ctx, s, at, payload, mode ?? 'all'),
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
		setSplit: (s, split, topLeft) => view.setSplit(ctx, s, split, topLeft),
		setComment: (s, at, text, author, threaded) =>
			annotations.setComment(ctx, s, at, text, author, threaded),
		setCommentReply: (s, at, index, text) => annotations.setCommentReply(ctx, s, at, index, text),
		setCommentResolved: (s, at, resolved) => annotations.setCommentResolved(ctx, s, at, resolved),
		setHyperlink: (s, range, link) => annotations.setHyperlink(ctx, s, range, link),
		addConditionalFormat: (s, cf) => {
			annotations.addConditionalFormat(ctx, s, cf);
			if (cf.ranges.length) prepareConditionalFormat(cf);
		},
		clearConditionalFormats: (s, range) => annotations.clearConditionalFormats(ctx, s, range),
		setDataValidation: (s, dv, range) => annotations.setDataValidation(ctx, s, dv, range),
		setAutoFilter: (s, range) => filter.setAutoFilter(ctx, s, range),
		filterColumn: (s, col, vals) => filter.filterColumn(ctx, s, col, vals),
		setFilterCriteria: (s, col, criteria, tableName) =>
			filter.setColumnCriteria(ctx, s, col, criteria, tableName),
		reapplyFilter: (s, at) => filter.reapplyFilter(ctx, s, at),
		applyAdvancedFilter: (s, options) => advancedFilter.applyAdvancedFilter(ctx, s, options),
		sortByColumn: (s, col, descending) => filter.sortByColumn(ctx, s, col, descending ?? false),
		createTable: (s, range, hasHeader, styleName) =>
			tables.createTable(ctx, s, range, hasHeader, styleName),
		setSheetView: (s, patch) => view.setSheetView(ctx, s, patch),
		setIgnoredErrors: (s, list) => view.setIgnoredErrors(ctx, s, list),
		setDefinedName: (name) => view.setDefinedName(ctx, name),
		deleteDefinedName: (name, localSheet) => view.deleteDefinedName(ctx, name, localSheet),
		editPictureEffects: (s, index, patch) => drawings.editPictureEffects(ctx, s, index, patch),
		setPictureCrop: (s, index, crop) => drawings.setPictureCrop(ctx, s, index, crop),
		addShape: (s, anchor, value) => nativeShapes.addShape(ctx, s, anchor, value),
		setTextBoxText: (s, index, text) => nativeShapes.setTextBoxText(ctx, s, index, text),
		setShapeEquation: (s, index, equation, value) =>
			nativeShapes.setShapeEquation(ctx, s, index, equation, value),
		reorderDrawing: (s, index, direction) => arrange.reorderDrawing(ctx, s, index, direction),
		groupDrawings: (s, indices, name) => groups.groupDrawings(ctx, s, indices, name),
		ungroupDrawing: (s, index) => groups.ungroupDrawing(ctx, s, index),
		alignDrawings: (s, indices, alignment) => arrange.alignDrawings(ctx, s, indices, alignment),
		flipDrawing: (s, index, axis) => arrange.flipDrawing(ctx, s, index, axis),
		setDrawingHidden: (s, index, hidden) => arrange.setDrawingHidden(ctx, s, index, hidden),
		rotateDrawing: (s, index, degrees) => arrange.rotateDrawing(ctx, s, index, degrees),
		setDrawingAnchor: (s, index, anchor) => drawings.setDrawingAnchor(ctx, s, index, anchor),
		createPivotTable: (s, destination, options) =>
			pivots.createPivotTable(ctx, s, destination, options),
		refreshPivotTable: (s, name, patch) => refreshPivotTable(ctx, s, name, patch),
		refreshAllPivotTables: () =>
			session.batch('Refresh All PivotTables', () => refreshAllPivotTables(ctx)),
		addPivotChart: (s, pivotSheet, name, anchor) =>
			pivotCharts.addPivotChart(ctx, s, pivotSheet, name, anchor),
		showPivotChartButtons: (s, index, visible) =>
			pivotCharts.showPivotChartButtons(ctx, s, index, visible),
		deletePivotTable: (s, name) => pivots.deletePivotTable(ctx, s, name),
		setSlicerSelection: (cache, keys) => slicers.setSlicerSelection(ctx, cache, keys),
		updateSlicer: (name, patch) => slicers.updateSlicer(ctx, name, patch),
		insertSlicer: (s, source) => slicers.insertSlicer(ctx, s, source),
		deleteSlicer: (name) => slicers.deleteSlicer(ctx, name),
		setTimelineSelection: (cache, selection) =>
			timelines.setTimelineSelection(ctx, cache, selection),
		updateTimeline: (name, patch) => timelines.updateTimeline(ctx, name, patch),
		insertTimeline: (s, source) => timelines.insertTimeline(ctx, s, source),
		deleteTimeline: (name) => timelines.deleteTimeline(ctx, name),
		deleteDrawing: (s, index) => drawings.deleteDrawing(ctx, s, index),
		addWordArt: (s, anchor, value) => drawings.addWordArt(ctx, s, anchor, value),
		insertLocalObject: (s, input) => insertLocalObject(ctx, s, input),
		editWordArt: (s, index, value) => drawings.editWordArt(ctx, s, index, value),
		editSmartArt: (s, index, edit) => smartArt.editSmartArt(ctx, s, index, edit),
		formatSmartArt: (s, index, change) => smartArt.formatSmartArt(ctx, s, index, change),
		changeSmartArtLayout: (s, index, change) =>
			smartArt.changeSmartArtLayout(ctx, s, index, change),
		insertSmartArt: (s, insertion) => smartArt.insertSmartArt(ctx, s, insertion),
		createSparklineGroup: (s, source, locations, patch) =>
			sparklines.createSparklineGroup(ctx, s, source, locations, patch),
		setSparklineGroups: (s, ranges, patch) => sparklines.setSparklineGroups(ctx, s, ranges, patch),
		setSparklineDateRange: (s, ranges, source) =>
			sparklines.setSparklineDateRange(ctx, s, ranges, source),
		groupSparklines: (s, ranges) => sparklineGrouping.groupSparklines(ctx, s, ranges),
		ungroupSparklines: (s, ranges) => sparklineGrouping.ungroupSparklines(ctx, s, ranges),
		deleteSparklines: (s, ranges, wholeGroups) =>
			sparklines.deleteSparklines(ctx, s, ranges, wholeGroups),
		addChart: (s, chart) => charts.addChart(ctx, s, chart),
		updateChart: (s, index, patch) => charts.updateChart(ctx, s, index, patch),
		editChart: (s, index, intent) => charts.editChart(ctx, s, index, intent),
		editChartEx: (s, index, operation) => chartEx.editChartEx(ctx, s, index, operation),
		addChartEx: (s, chart) => chartEx.addChartEx(ctx, s, chart),
		addImage: (s, bytes, type, anchor, name, svg) =>
			charts.addImage(ctx, s, bytes, type, anchor, name, svg),
		addCellPicture: (s, at, bytes, text) => pictures.addCellPicture(ctx, s, at, bytes, text),
		placePictureInCell: (s, index) => pictures.placePictureInCell(ctx, s, index),
		placePictureOverCells: (s, at) => pictures.placePictureOverCells(ctx, s, at),
		setCellPictureAltText: (s, at, text) => pictures.setCellPictureAltText(ctx, s, at, text),
		replaceConditionalFormat: (s, index, cf) => {
			cfEdits.replaceConditionalFormat(ctx, s, index, cf);
			prepareConditionalFormat(cf);
		},
		removeConditionalFormat: (s, index) => cfEdits.removeConditionalFormat(ctx, s, index),
		setConditionalRulePriority: (s, f, r, priority) =>
			cfEdits.setConditionalRulePriority(ctx, s, f, r, priority),
		moveConditionalRule: (s, f, r, direction) =>
			cfEdits.moveConditionalRule(ctx, s, f, r, direction),
		updateTable: (s, table, patch) => tableEdits.updateTable(ctx, s, table, patch),
		resizeTable: (s, table, range) => tableEdits.resizeTable(ctx, s, table, range),
		convertTableToRange: (s, table) => tableEdits.convertTableToRange(ctx, s, table),
		applyCellStyle: (s, ranges, name) => cellStyle.applyCellStyle(ctx, s, ranges, name),
		setSheetBackground: (s, image) => setSheetBackground(ctx, s, image),
		setPageSetup: (s, patch) => page.setPageSetup(ctx, s, patch),
		setPageBreak: (s, axis, before, present) =>
			pageBreaks.setPageBreak(ctx, s, axis, before, present),
		resetPageBreaks: (s) => pageBreaks.resetPageBreaks(ctx, s),
		setPrintTitles: (s, patch) => printTitles.setPrintTitles(ctx, s, patch),
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
		applySubtotal: (s, range, options) => subtotal.applySubtotal(ctx, s, range, options),
		removeSubtotals: (s, range, hasHeader) => subtotal.removeSubtotals(ctx, s, range, hasHeader),
		consolidate: (s, destination, sources, options) =>
			consolidate.consolidateRanges(ctx, s, destination, sources, options),
		goalSeek: (request) => goalSeek.goalSeek(ctx, request),
		createDataTable: (s, request) => dataTable.createDataTable(ctx, s, request),
		addScenario: (s, draft) => scenarios.addScenario(ctx, s, draft),
		editScenario: (s, i, draft) => scenarios.editScenario(ctx, s, i, draft),
		deleteScenario: (s, i) => scenarios.deleteScenario(ctx, s, i),
		showScenario: (s, i) => scenarios.showScenario(ctx, s, i),
		mergeScenarios: (s, from) => mergeScenarios(ctx, s, from),
		addCustomView: (draft) => customViews.addCustomView(ctx, draft),
		showCustomView: (guid) => customViews.showCustomView(ctx, guid),
		deleteCustomView: (guid) => customViews.deleteCustomView(ctx, guid),
		setCalcMode: (mode) => calcMode.setCalcMode(ctx, mode),
		setDocumentProperties: (patch) => docProps.setDocumentProperties(ctx, patch),
		setTheme: (patch, label) => themeEdits.setTheme(ctx, patch, label),
		setAutoRecalc: (enabled) => calcMode.setCalcMode(ctx, enabled ? 'auto' : 'manual'),
		autoRecalc: () => autoRecalc && calcMode.calcModeOf(workbook) === 'auto',
		calculateNow(options) {
			calculator.now(options?.full ?? false);
			emit({ kind: 'cells', label: 'Calculate now', structural: false });
		},
		prepareCalculation: (opts) => calculator.prepare(opts?.timeRemaining),
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
		applyExternal: (apply) => applyExternal({ calculator, emit, busy: () => !!open }, apply),
	};
	return session;
}
