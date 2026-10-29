import type { ChartExSpace } from '../../chart/ex/model';
import type { DocumentPropertiesPatch } from './doc-properties';
import type { DataTableRequest } from './data-table';
import type { GoalSeekRequest, GoalSeekResult } from './goal-seek';
import type { ScenarioDraft } from './scenarios';
import type { ExternalChange } from './external';
import type { CellAddress, CellRange } from '../address';
import type {
	BorderEdge,
	CellStyle,
	CellValue,
	ChartObject,
	Color,
	PivotTableOptions,
	PivotTablePatch,
	ConditionalFormat,
	ImageObject,
	DataValidation,
	DefinedName,
	DrawingAnchor,
	FreezePane,
	Hyperlink,
	PageSetup,
	PrintOptions,
	SheetProtection,
	SheetState,
	SheetView,
	Table,
	Workbook,
} from '../model';
import type { FontView } from '../layout/types';
import type { StylePatch } from '../styles';
import type { CalcEngine } from './deps';
import type { ChartPatch } from './charts';
import type { SmartArtEditResult, SmartArtInsertion, SmartArtLayoutChange } from './smart-art';
import type { DiagramNodeEdit } from '../../diagram/edit/node-edit';
import type { DiagramFormatChange } from '../../diagram/edit/format';
import type { ChartEditIntent } from '../../chart/edit/intent';
import type { TablePatch, TableRef } from './table-edits';
import type { SparklineGroupPatch } from './sparklines';

/** Border presets of the ribbon's border drop-down. */
export type BorderPreset =
	| 'all'
	| 'outside'
	| 'thickOutside'
	| 'none'
	| 'top'
	| 'bottom'
	| 'left'
	| 'right'
	| 'inside'
	| 'insideH'
	| 'insideV'
	| 'thickBottom'
	| 'doubleBottom'
	| 'topAndBottom';

/** What `clearRange` removes. */
export type ClearWhat = 'all' | 'contents' | 'formats' | 'comments' | 'hyperlinks';

export type PasteMode =
	| 'all'
	| 'values'
	| 'formats'
	| 'formulas'
	| 'noBorders'
	| 'mergeFormats'
	| 'widths'
	| 'comments'
	| 'validation'
	| 'transpose';
export type PasteOperation = 'none' | 'add' | 'subtract' | 'multiply' | 'divide';

/** Paste content and independent modifiers. Legacy string modes remain supported. */
export interface PasteOptions {
	mode?: Exclude<PasteMode, 'transpose'>;
	transpose?: boolean;
	skipBlanks?: boolean;
	operation?: PasteOperation;
}
export type PasteRequest = PasteMode | PasteOptions;

/** AutoFill infers a series; Fill Down / Right repeat cells with translated formulas. */
export type FillMode = 'auto' | 'copy';

export type MergeMode = 'merge' | 'center' | 'across';

export type {
	ClipboardCell,
	ClipboardCells,
	ClipboardPayload,
	FindMatch,
	FindQuery,
} from './types-clipboard';
import type { TimelineSelection } from '../slicer/timeline-periods';
import type { SlicerPatch, SlicerSourceRef } from './slicers';
import type { TimelinePatch } from './timelines';
import type { ClipboardPayload, FindMatch, FindQuery } from './types-clipboard';

export type WorkbookChangeKind =
	| 'cells'
	| 'format'
	| 'structure'
	| 'sheets'
	| 'annotations'
	| 'view'
	| 'names'
	| 'batch'
	| 'undo'
	| 'redo'
	/** A collaborator's change applied through `applyExternal`. */
	| 'remote';

export interface WorkbookChange {
	kind: WorkbookChangeKind;
	/** Undo label of the edit. */
	label: string;
	/** The sheet the edit touched, when it touched one. */
	sheet?: number;
	ranges?: CellRange[];
	/** Rows, columns or sheets moved: views must rebuild metrics, not only repaint cells. */
	structural: boolean;
	/** Applied through `applyExternal` (a collaborator or a shared undo), not recorded in history. */
	external?: true;
}

export interface ValidationFailure {
	ok: false;
	style: 'stop' | 'warning' | 'information';
	title?: string;
	message: string;
}

export type ValidationResult = { ok: true } | ValidationFailure;

export interface EditSessionOptions {
	/** Recalculate formulas after every edit (default true). */
	recalc?: boolean;
	/** Undo depth (default 100). */
	historyLimit?: number;
	/** Base of new sheet names (`Sheet` gives `Sheet2`, ...), for localized UIs. */
	defaultSheetBase?: string;
	/**
	 * Re-fit rows without a custom height after edits that change cell text or formats, as Excel
	 * grows and shrinks such rows (default true). The re-fit undoes together with the edit.
	 */
	autoRowHeight?: boolean;
	/** Text measurement for the row re-fit, in CSS px (default: an estimate from the font size). */
	measureText?: (text: string, font: FontView) => number;
}

/** The editing API every UI drives. Each method is one undo step; `batch` groups several. */
export interface EditSession {
	readonly workbook: Workbook;
	readonly calc: CalcEngine;
	setCellInput(sheet: number, row: number, col: number, text: string): void;
	setCellValue(sheet: number, row: number, col: number, value: CellValue): void;
	setRangeValues(sheet: number, at: CellAddress, values: CellValue[][]): void;
	clearRange(sheet: number, range: CellRange, what: ClearWhat): void;
	applyStyle(sheet: number, ranges: CellRange[], patch: StylePatch): void;
	setBorders(sheet: number, range: CellRange, preset: BorderPreset, edge?: BorderEdge): void;
	insertRows(sheet: number, at: number, count: number): void;
	deleteRows(sheet: number, at: number, count: number): void;
	insertColumns(sheet: number, at: number, count: number): void;
	deleteColumns(sheet: number, at: number, count: number): void;
	insertCellsShift(sheet: number, range: CellRange, direction: 'down' | 'right'): void;
	deleteCellsShift(sheet: number, range: CellRange, direction: 'up' | 'left'): void;
	setColumnWidth(
		sheet: number,
		cols: number[],
		width: number | 'auto',
		measure?: (text: string, style: CellStyle) => number,
	): void;
	setRowHeight(sheet: number, rows: number[], height: number | 'auto'): void;
	setHidden(sheet: number, axis: 'row' | 'col', indices: number[], hidden: boolean): void;
	merge(sheet: number, range: CellRange, mode: MergeMode): void;
	unmerge(sheet: number, range: CellRange): void;
	sort(
		sheet: number,
		range: CellRange,
		keys: { col: number; descending?: boolean }[],
		hasHeader: boolean,
	): void;
	fill(sheet: number, source: CellRange, target: CellRange, mode?: FillMode): void;
	copy(sheet: number, range: CellRange): ClipboardPayload;
	cut(sheet: number, range: CellRange): ClipboardPayload;
	paste(
		sheet: number,
		at: CellAddress | CellRange,
		payload: ClipboardPayload | string,
		mode?: PasteRequest,
	): CellRange;
	findAll(query: FindQuery): FindMatch[];
	replaceAll(query: FindQuery, replacement: string): number;
	/** Replaces the query in one matched cell; false when the cell no longer matches. */
	replaceOne(query: FindQuery, replacement: string, at: FindMatch): boolean;
	addSheet(name?: string, at?: number): number;
	deleteSheet(index: number): void;
	renameSheet(index: number, name: string): void;
	moveSheet(from: number, to: number): void;
	duplicateSheet(index: number): number;
	setSheetState(index: number, state: SheetState): void;
	setTabColor(index: number, color?: Color): void;
	setFreeze(sheet: number, freeze: FreezePane | undefined): void;
	setComment(sheet: number, at: CellAddress, text: string | undefined, author: string): void;
	setHyperlink(sheet: number, range: CellRange, link: Omit<Hyperlink, 'range'> | undefined): void;
	addConditionalFormat(sheet: number, format: ConditionalFormat): void;
	clearConditionalFormats(sheet: number, range?: CellRange): void;
	setDataValidation(sheet: number, validation: DataValidation | undefined, range: CellRange): void;
	setAutoFilter(sheet: number, range: CellRange | undefined): void;
	/** Filters an auto-filter column to display texts in `values` (`''` keeps blanks); undefined clears. */
	filterColumn(sheet: number, col: number, values: string[] | undefined): void;
	/** Sorts the auto-filter (or table) containing `col` by that column, keeping the header. */
	sortByColumn(sheet: number, col: number, descending?: boolean): void;
	createTable(sheet: number, range: CellRange, hasHeader: boolean, styleName?: string): Table;
	setSheetView(sheet: number, patch: Partial<SheetView>): void;
	setDefinedName(name: DefinedName): void;
	deleteDefinedName(name: string, localSheet?: number): void;
	/** Moves or resizes the picture, chart or shape at `index` of the sheet's drawings. */
	/** Sets source crop insets on an image, with history. */
	/** Applies admitted picture effects while retaining opaque operations. */
	editPictureEffects(
		sheet: number,
		index: number,
		patch: import('../../drawingml/picture-effects').PictureEffectPatch,
	): void;
	setPictureCrop(
		sheet: number,
		index: number,
		crop: import('../../drawingml/picture-crop').DrawingPictureCrop,
	): void;
	setDrawingAnchor(sheet: number, index: number, anchor: DrawingAnchor): void;
	/** Deletes the picture, chart or shape at `index` of the sheet's drawings. */
	deleteDrawing(sheet: number, index: number): void;
	/** Applies a SmartArt Design or text pane command to the SmartArt at `index` (one undo step). */
	editSmartArt(sheet: number, index: number, edit: DiagramNodeEdit): SmartArtEditResult;
	formatSmartArt(sheet: number, index: number, change: DiagramFormatChange): boolean;
	/** SmartArt Design > Change Layout; false when the diagram already uses the layout. */
	changeSmartArtLayout(sheet: number, index: number, change: SmartArtLayoutChange): boolean;
	/** Insert > SmartArt; returns the new drawing's index. */
	insertSmartArt(sheet: number, insertion: SmartArtInsertion): number;
	/** Creates a worksheet-range PivotTable and its cached output. */
	createPivotTable(sheet: number, destination: CellAddress, options: PivotTableOptions): string;
	/** Refreshes cached records and output; an optional patch reconfigures fields. */
	refreshPivotTable(sheet: number, name: string, patch?: PivotTablePatch): void;
	/** Removes the PivotTable and its output cells as one undo step. */
	deletePivotTable(sheet: number, name: string): void;
	/** Filters a slicer's table column or PivotTable field to the item keys (null: all items). */
	setSlicerSelection(cache: string, keys: readonly string[] | null): void;
	/** Changes a slicer's caption, style, columns, button height or sort order. */
	updateSlicer(name: string, patch: SlicerPatch): void;
	/** Inserts a slicer for a table column or PivotTable field; returns the slicer's name. */
	insertSlicer(sheet: number, source: SlicerSourceRef): string;
	/** Deletes a slicer (and its cache when unused). */
	deleteSlicer(name: string): void;
	/** Filters a timeline's PivotTable to an ISO date range (null: every date). */
	setTimelineSelection(cache: string, selection: TimelineSelection | null): void;
	/** Changes a timeline's caption, time level or visible parts. */
	updateTimeline(name: string, patch: TimelinePatch): void;
	/** Inserts a timeline for a date field of a PivotTable; returns the timeline's name. */
	insertTimeline(sheet: number, source: { pivot: string; field: string }): string;
	/** Deletes a timeline (and its cache). */
	deleteTimeline(name: string): void;
	/** Creates a group, mapping data rows or columns to a one-dimensional location range. */
	createSparklineGroup(
		sheet: number,
		source: string,
		locations: CellRange,
		patch?: SparklineGroupPatch,
	): void;
	/** Settings affect each group containing a selected sparkline. */
	setSparklineGroups(sheet: number, ranges: CellRange[], patch: SparklineGroupPatch): void;
	/** A continuous date vector, or empty to return to a general axis. */
	setSparklineDateRange(sheet: number, ranges: CellRange[], source?: string): void;
	/** Clears selected hosts, or their whole groups. Cell contents remain. */
	deleteSparklines(sheet: number, ranges: CellRange[], wholeGroups?: boolean): void;
	/** Adds a chart drawn from the model; returns its index in the sheet's drawings. */
	addChart(sheet: number, chart: Omit<ChartObject, 'kind'>): number;
	/** Edits the chart at `index` of the sheet's drawings through its summary fields. */
	updateChart(sheet: number, index: number, patch: ChartPatch): void;
	/** Applies a `chart/edit` intent (a Format pane edit) to the chart at `index`. */
	editChart(sheet: number, index: number, intent: ChartEditIntent): void;
	/** Extended chart operation with atomic source-cell synchronization and history. */
	editChartEx(sheet: number, index: number, operation: (space: ChartExSpace) => ChartExSpace): void;
	/** Inserts a picture (bytes stored as a new media part) and returns it. */
	addImage(
		sheet: number,
		bytes: Uint8Array,
		contentType: string,
		anchor: DrawingAnchor,
		name?: string,
	): ImageObject;
	/** Replaces the conditional format at `index` (Rules Manager > Edit Rule). */
	replaceConditionalFormat(sheet: number, index: number, format: ConditionalFormat): void;
	/** Removes the conditional format at `index` (Rules Manager > Delete Rule). */
	removeConditionalFormat(sheet: number, index: number): void;
	/** Gives a rule a new priority (1 is evaluated first); every rule is renumbered 1..n. */
	setConditionalRulePriority(
		sheet: number,
		formatIndex: number,
		ruleIndex: number,
		priority: number,
	): void;
	/** Moves a rule one step up or down in the evaluation order. */
	moveConditionalRule(
		sheet: number,
		formatIndex: number,
		ruleIndex: number,
		direction: 'up' | 'down',
	): void;
	/**
	 * Updates a table (by index in `sheet.tables` or by name): style options, name (structured
	 * references follow), header and totals rows, columns and range.
	 */
	updateTable(sheet: number, table: TableRef, patch: TablePatch): void;
	/** Resizes a table; its header row must stay on the same row. */
	resizeTable(sheet: number, table: TableRef, range: CellRange): void;
	/** Converts a table to a plain range; structured references become absolute A1 references. */
	convertTableToRange(sheet: number, table: TableRef): void;
	/** Applies a named cell style (built-in such as `Good`, or one the workbook defines). */
	applyCellStyle(sheet: number, ranges: CellRange[], name: string): void;
	/** Patches the page setup; `undefined` entries clear a setting. */
	setPageSetup(sheet: number, patch: { [K in keyof PageSetup]?: PageSetup[K] | undefined }): void;
	setPageBreak(sheet: number, axis: 'row' | 'col', before: number, present?: boolean): void;
	resetPageBreaks(sheet: number): void;
	setPrintTitles(
		sheet: number,
		patch: {
			rows?: import('../layout/print-titles').PrintTitleSpan | undefined;
			columns?: import('../layout/print-titles').PrintTitleSpan | undefined;
		},
	): void;
	/** Patches the print options (gridlines, headings, centring); false clears one. */
	setPrintOptions(sheet: number, patch: Partial<PrintOptions>): void;
	/**
	 * Protects a sheet (`undefined` unprotects). A non-empty `password` stores its legacy hash and
	 * drops the modern one, `''` removes both; unprotecting with a wrong password throws.
	 */
	setSheetProtection(
		sheet: number,
		protection: SheetProtection | undefined,
		password?: string,
	): void;
	/** Locks or unlocks the workbook structure, optionally with a password (checked on unlock). */
	setWorkbookProtection(locked: boolean, password?: string): void;
	/** Raises the outline level of rows `from`..`to` (inclusive). */
	groupRows(sheet: number, from: number, to: number): void;
	ungroupRows(sheet: number, from: number, to: number): void;
	/** Raises the outline level of columns `from`..`to` (inclusive). */
	groupColumns(sheet: number, from: number, to: number): void;
	ungroupColumns(sheet: number, from: number, to: number): void;
	/** The sheet's standard column width in characters; undefined restores Excel's default. */
	setDefaultColumnWidth(sheet: number, width: number | undefined): void;
	/** Removes rows whose `columns` (absolute; empty = all) repeat an earlier row's display text. */
	removeDuplicates(
		sheet: number,
		range: CellRange,
		columns: number[],
		hasHeader: boolean,
	): { removed: number; remaining: number };
	/**
	 * Goal Seek: changes `changing` until `target` (a formula) equals `goal`, honouring the
	 * workbook's iteration settings. One undo step; undo it to cancel. Throws when the request is
	 * not valid (see `checkGoalSeek`).
	 */
	goalSeek(request: GoalSeekRequest): GoalSeekResult;
	/** Turns a selection into a what-if data table `{=TABLE()}` and computes it. */
	createDataTable(sheet: number, request: DataTableRequest): void;
	/** Scenario Manager: adds a scenario recording the changing cells; returns its index. */
	addScenario(sheet: number, draft: ScenarioDraft): number;
	editScenario(sheet: number, index: number, draft: ScenarioDraft): void;
	deleteScenario(sheet: number, index: number): void;
	/** Writes a scenario's values into its changing cells. */
	showScenario(sheet: number, index: number): void;
	/** Appends the scenarios of another sheet (renaming taken names); returns how many were added. */
	mergeScenarios(sheet: number, from: number): number;
	/** Automatic or manual calculation (saved with the workbook). */
	setCalcMode(mode: 'auto' | 'manual'): void;
	/**
	 * Changes document properties (File > Info): a value sets a core, extended or custom field,
	 * `null` clears it; `custom` replaces the whole list. One undo step; returns whether anything
	 * changed.
	 */
	setDocumentProperties(patch: DocumentPropertiesPatch): boolean;
	/** `setCalcMode(enabled ? 'auto' : 'manual')`. */
	setAutoRecalc(enabled: boolean): void;
	/**
	 * Whether edits recalculate: false when the session was created with `recalc: false` (which
	 * always wins) or the workbook is in manual calculation mode.
	 */
	autoRecalc(): boolean;
	/**
	 * Calculate Now (F9), not an undo step: recalculates the formulas changed since the last
	 * calculation (edits made in manual mode, structural edits) with their dependents and the
	 * volatile functions. The first calculation of a session, a change of the iteration settings
	 * and structural edits the graph could not follow recalculate everything. `full` always does
	 * (Ctrl+Alt+F9: every formula, with the dependency graph rebuilt).
	 */
	calculateNow(options?: { full?: boolean }): void;
	/**
	 * Builds the formula dependency graph ahead of the first edit, so that edit does not wait
	 * for every formula to be parsed. Call it repeatedly from idle time: each call works until
	 * `timeRemaining()` (milliseconds left, as `requestIdleCallback`'s deadline reports them)
	 * reaches zero and returns true once the graph is ready. Without `timeRemaining` it finishes
	 * in one call. Edits made before it is done stay correct: they finish the remaining work
	 * first. A no-op (true) when edits do not recalculate.
	 */
	prepareCalculation(options?: { timeRemaining?: () => number }): boolean;
	/**
	 * Recalculates the formulas of one sheet and their dependents (Shift+F9). After structural
	 * edits made in manual mode that the dependency graph could not follow (sheets added, deleted
	 * or moved, cells inserted or deleted, ...) the whole workbook is recalculated instead.
	 */
	calculateSheet(sheet: number): void;
	/** Validates `value` against the cell's data validation, using this session's calc engine. */
	validate(sheet: number, row: number, col: number, value: CellValue): ValidationResult;
	batch(label: string, fn: () => void): void;
	undo(): boolean;
	redo(): boolean;
	canUndo(): boolean;
	canRedo(): boolean;
	undoLabel(): string | undefined;
	redoLabel(): string | undefined;
	onChange(listener: (change: WorkbookChange) => void): () => void;
	/**
	 * Applies a change that did not start here (a collaborator's edit, a shared undo): `apply`
	 * mutates the workbook and describes what it touched (or returns undefined when nothing
	 * changed). It is recalculated and announced to listeners with `external: true`, but not
	 * recorded in the undo history, so a session bound to a shared document undoes through its
	 * binding. Throws inside a batch.
	 */
	applyExternal(apply: () => ExternalChange | undefined): boolean;
}
