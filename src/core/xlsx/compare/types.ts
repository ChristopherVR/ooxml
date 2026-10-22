import type { CellAddress, CellRange } from '../address';
import type {
	CellStyle,
	CellValue,
	ConditionalFormat,
	DataValidation,
	DefinedName,
	SheetState,
	Table,
} from '../model';

/** What a comparison can look at. Cell categories follow Excel's Spreadsheet Compare. */
export type CompareCategory =
	/** A constant typed into a cell (not a formula, not a spilled or array result). */
	| 'enteredValue'
	/** A formula added, removed or edited (data table results show their `TABLE()` formula). */
	| 'formula'
	/** The cached result of an unchanged formula, a spilled value or an array formula result. */
	| 'calculatedValue'
	/** The effective format: number format, font, fill, border, alignment, protection. */
	| 'format'
	/** Sheets added, removed, renamed, moved, shown or hidden. */
	| 'sheet'
	/** Defined names added, removed or changed. */
	| 'name'
	| 'merge'
	| 'table'
	| 'dataValidation'
	| 'conditionalFormat';

export const COMPARE_CATEGORIES: readonly CompareCategory[] = [
	'enteredValue',
	'formula',
	'calculatedValue',
	'format',
	'sheet',
	'name',
	'merge',
	'table',
	'dataValidation',
	'conditionalFormat',
];

/** Parts of a workbook no comparison looks at; listed in every result so a report stays honest. */
export const NOT_COMPARED = [
	'rowAndColumnLayout',
	'rowAndColumnFormats',
	'comments',
	'hyperlinks',
	'drawings',
	'pivotTables',
	'conditionalFormatExtensions',
	'sparklines',
	'slicers',
	'autoFilter',
	'sheetViews',
	'pageSetup',
	'sheetProtection',
	'documentProperties',
	'theme',
	'workbookSettings',
	'macrosAndUnmodelledParts',
] as const;
export type UncomparedFeature = (typeof NOT_COMPARED)[number];

/**
 * How sheets were paired. `auto` (the default) pairs sheets with the same name first, then
 * sheets with the same `sheetId` (a rename), so a file whose producer renumbered the ids still
 * pairs by name. `name` and `sheetId` use one rule only.
 */
export type SheetMatch = 'auto' | 'name' | 'sheetId';

export interface CompareOptions {
	/** Categories to compare; default all of {@link COMPARE_CATEGORIES}. */
	categories?: readonly CompareCategory[];
	/** Stop after this many differences and set `truncated`; default unlimited. */
	maxDifferences?: number;
	sheetMatch?: SheetMatch;
}

/** The parts of a format compared by {@link CompareCategory} `format`. */
export type FormatAttribute =
	| 'numberFormat'
	| 'font'
	| 'fill'
	| 'border'
	| 'alignment'
	| 'protection'
	/** The fonts of the runs of a rich-text string. */
	| 'richText';

/** One side of a cell difference; an empty cell is `{ value: null }`. */
export interface CellSide {
	value: CellValue;
	/** Without the leading `=`. */
	formula?: string;
	arrayRange?: CellRange;
	/** The formula is evaluated with implicit intersection (legacy semantics), where that matters. */
	legacyFormula?: true;
	/** The effective format, for `format` differences. */
	style?: CellStyle;
}

export interface CellDifference {
	kind: 'enteredValue' | 'formula' | 'calculatedValue' | 'format';
	address: CellAddress;
	before: CellSide;
	after: CellSide;
	/** What changed, for `format` differences. */
	attributes?: FormatAttribute[];
}

/** A rectangle of neighbouring cells with the same kind of difference. */
export interface CellDifferenceRange {
	kind: CellDifference['kind'];
	range: CellRange;
	/** Cells in the range (every cell of it differs). */
	count: number;
	/** The changed attributes shared by the range, for `format` ranges. */
	attributes?: FormatAttribute[];
}

/** A difference in the structure of the workbook (sheets and defined names). */
export type WorkbookDifference =
	| { kind: 'sheetAdded'; sheet: string; index: number }
	| { kind: 'sheetRemoved'; sheet: string; index: number }
	| { kind: 'sheetRenamed'; from: string; to: string }
	| { kind: 'sheetMoved'; sheet: string; from: number; to: number }
	| { kind: 'sheetState'; sheet: string; from: SheetState; to: SheetState }
	| { kind: 'nameAdded'; name: string; scope?: string; after: DefinedName }
	| { kind: 'nameRemoved'; name: string; scope?: string; before: DefinedName }
	| {
			kind: 'nameChanged';
			name: string;
			scope?: string;
			before: DefinedName;
			after: DefinedName;
	  };

/** A difference in a sheet's merges, tables, data validation or conditional formats. */
export type SheetDifference =
	| { kind: 'mergeAdded' | 'mergeRemoved'; range: CellRange }
	| { kind: 'tableAdded'; table: string; after: Table }
	| { kind: 'tableRemoved'; table: string; before: Table }
	| {
			kind: 'tableChanged';
			table: string;
			before: Table;
			after: Table;
			fields: ('range' | 'columns' | 'headerRow' | 'totalsRow' | 'style')[];
	  }
	| { kind: 'dataValidationAdded'; after: DataValidation }
	| { kind: 'dataValidationRemoved'; before: DataValidation }
	/** Conditional format differences are per rule: each is a format with that one rule. */
	| { kind: 'conditionalFormatAdded'; after: ConditionalFormat }
	| { kind: 'conditionalFormatRemoved'; before: ConditionalFormat }
	/** The rule moved above or below other rules (priority order, not numbering). */
	| { kind: 'conditionalFormatReordered'; before: ConditionalFormat; after: ConditionalFormat };

/** A pair of sheets present in both workbooks and what differs between them. */
export interface SheetComparison {
	/** The name in the after workbook. */
	name: string;
	beforeName: string;
	/** Zero-based positions in each workbook. */
	beforeIndex: number;
	afterIndex: number;
	/** Which rule paired the two sheets. */
	matchedBy: 'name' | 'sheetId';
	structure: SheetDifference[];
	/** Cell differences in row-major order; for one cell, `formula` and value kinds before `format`. */
	cells: CellDifference[];
	/** {@link cells} grouped into rectangles, in row-major order of their top-left cell. */
	ranges: CellDifferenceRange[];
}

export interface WorkbookComparison {
	workbook: WorkbookDifference[];
	/** Every paired sheet, in the after workbook's order. Added and removed sheets are in `workbook`. */
	sheets: SheetComparison[];
	/** Differences reported, cells and structure together. */
	count: number;
	/** `maxDifferences` was reached: more differences may exist. */
	truncated: boolean;
	categories: CompareCategory[];
	/** Categories left out by the options, then everything no comparison looks at. */
	notCompared: (CompareCategory | UncomparedFeature)[];
}
