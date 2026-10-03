import type { DocumentPropertiesPatch } from './doc-properties.js';
import type { CellAddress, CellRange } from '../address.js';
import type {
	BorderEdge,
	CellStyle,
	CellValue,
	ChartObject,
	Color,
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
} from '../model.js';
import type { FontView } from '../layout/types.js';
import type { StylePatch } from '../styles.js';
import type { CalcEngine } from './deps.js';
import type { ChartPatch } from './charts.js';
import type { TablePatch, TableRef } from './table-edits.js';

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

export type PasteMode = 'all' | 'values' | 'formats' | 'formulas' | 'transpose';

export type MergeMode = 'merge' | 'center' | 'across';

export type {
	ClipboardCell,
	ClipboardCells,
	ClipboardPayload,
	FindMatch,
	FindQuery,
} from './types-clipboard.js';
import type { ClipboardPayload, FindMatch, FindQuery } from './types-clipboard.js';

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
	| 'redo';

export interface WorkbookChange {
	kind: WorkbookChangeKind;
	/** Undo label of the edit. */
	label: string;
	/** The sheet the edit touched, when it touched one. */
	sheet?: number;
	ranges?: CellRange[];
	/** Rows, columns or sheets moved: views must rebuild metrics, not only repaint cells. */
	structural: boolean;
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
	fill(sheet: number, source: CellRange, target: CellRange): void;
	copy(sheet: number, range: CellRange): ClipboardPayload;
	cut(sheet: number, range: CellRange): ClipboardPayload;
	paste(
		sheet: number,
		at: CellAddress,
		payload: ClipboardPayload | string,
		mode?: PasteMode,
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
	setDrawingAnchor(sheet: number, index: number, anchor: DrawingAnchor): void;
	/** Deletes the picture, chart or shape at `index` of the sheet's drawings. */
	deleteDrawing(sheet: number, index: number): void;
	/** Adds a chart drawn from the model; returns its index in the sheet's drawings. */
	addChart(sheet: number, chart: Omit<ChartObject, 'kind'>): number;
	/** Edits the chart at `index` of the sheet's drawings (a loaded chart part is patched on save). */
	updateChart(sheet: number, index: number, patch: ChartPatch): void;
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
	/** Recalculates every formula (F9); not an undo step. */
	calculateNow(): void;
	/**
	 * Recalculates the formulas of one sheet and their dependents (Shift+F9). After structural
	 * edits made in manual mode the whole workbook is recalculated instead.
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
}
