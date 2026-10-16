import type { CellRange } from './address';
import type { CellValue } from './model';

/**
 * A summary function: the Summarize Values By choices and the per-field subtotal functions.
 * `count` counts non-empty values, `countNums` only numbers (Excel's COUNTA and COUNT).
 */
export type PivotAggregate =
	| 'sum'
	| 'count'
	| 'average'
	| 'min'
	| 'max'
	| 'product'
	| 'countNums'
	| 'stdDev'
	| 'stdDevp'
	| 'var'
	| 'varp';
/**
 * The subtotals of one row field. Without an entry the field shows the automatic subtotal (a
 * Total row, or the group row's value when subtotals sit above); an empty list shows none; a
 * list shows one subtotal row per function, below the group, computed from the source values.
 */
export interface PivotFieldSubtotals {
	field: string;
	functions: PivotAggregate[];
}
/** The units a date field can be grouped by, finest first (Excel's Group dialog order). */
export type PivotDateUnit =
	| 'seconds'
	| 'minutes'
	| 'hours'
	| 'days'
	| 'months'
	| 'quarters'
	| 'years';
/**
 * A date group on a date field (Group Selection): one derived field per unit, named
 * `Years (Date)`, replacing the date field on the axes. `start` and `end` (serials) are
 * automatic when omitted; `days` is the number of days per group when grouping by days only.
 */
export interface PivotDateGroup {
	field: string;
	by: PivotDateUnit[];
	start?: number;
	end?: number;
	days?: number;
	/**
	 * What a file written by Excel in another language (or with renamed group items) calls the
	 * group fields and their items, per unit: the field name and the item labels in item order,
	 * including the `<start` and `>end` items. Units that use the English names have no entry.
	 */
	localized?: Partial<Record<PivotDateUnit, { name: string; labels: string[] }>>;
}
/**
 * The captions a PivotTable writes into its cells. Excel takes them from its UI language, so a
 * workbook saved by a localized Excel holds them in that language: the row and column header
 * (`rowHeaderCaption`, `colHeaderCaption`), the grand total (`grandTotalCaption`), the Values
 * header (`dataCaption`) and the words of subtotal rows, which only the cached cells record.
 * Templates use `{0}` for the item (`{0} Total`) or the measure (`Total {0}`).
 */
export interface PivotCaptions {
	rowLabels: string;
	columnLabels: string;
	grandTotal: string;
	values: string;
	/** The report filter item that selects everything. */
	all: string;
	/** The label of a blank item. */
	blank: string;
	/** A group's automatic subtotal row or column. */
	subtotal: string;
	/** A measure's grand total, or its total with several measures. */
	measureTotal: string;
	/** The suffix of a custom subtotal row (`East Sum`). */
	functions: Record<PivotAggregate, string>;
	/** The default name of a value field (`Sum of {0}`). */
	aggregates: Record<PivotAggregate, string>;
}
/** A display name Excel keeps for an item of a field (`item/@n`), such as Eastern for East. */
export interface PivotItemName {
	field: string;
	item: CellValue;
	name: string;
}
/** Captions that differ from English; the rest keep their English text. */
export type PivotCaptionOverrides = Partial<Omit<PivotCaptions, 'functions' | 'aggregates'>> & {
	functions?: Partial<Record<PivotAggregate, string>>;
	aggregates?: Partial<Record<PivotAggregate, string>>;
};
export type PivotValuesPosition = 'last' | 'between' | 'first' | number;
export type PivotLayout = 'tabular' | 'compact' | 'outline';
export interface PivotPageFilter {
	field: string;
	/** Omitted means All; null selects the cached blank item. */
	item?: CellValue;
}
export interface PivotValueField {
	field: string;
	aggregate: PivotAggregate;
	caption: string;
}
export interface PivotValueFieldOptions {
	field: string;
	aggregate?: PivotAggregate;
	caption?: string;
}

/** Items hidden from one field (what a slicer or the item filter list sets). */
export interface PivotItemFilter {
	field: string;
	/** Cached items that are not shown. At least one item stays visible. */
	hidden: CellValue[];
}

/** A number group on a field (Group Field: starting at, ending at, by); integers only. */
export interface PivotFieldGroup {
	field: string;
	start: number;
	end: number;
	by: number;
}

/** A date range filter (what a timeline sets): dates from `start` to `end` inclusive, as serials. */
export interface PivotDateFilter {
	field: string;
	start: number;
	end: number;
}

/** A worksheet-range pivot with any number of row and column fields and sixteen measures. */
export interface PivotTable {
	name: string;
	/** Qualified A1 source range, rewritten with the workbook's other references. */
	source: string;
	rowField: string;
	layout?: PivotLayout;
	subtotalPosition?: 'top' | 'bottom';
	secondRowField?: string;
	/** The third and following row fields, outermost first. Requires `secondRowField`. */
	moreRowFields?: string[];
	/** Automatic outer-field totals below each group. Defaults to true. */
	subtotals?: boolean;
	/** Per-field subtotal functions; overrides `subtotals` for the fields it names. */
	fieldSubtotals?: PivotFieldSubtotals[];
	/** Date groups: each adds derived fields (`Years (Date)`) usable as row fields. */
	dateGroups?: PivotDateGroup[];
	/** Repeat outer labels on detail rows. Defaults to false. */
	repeatLabels?: boolean;
	columnField?: string;
	/** The second and following column fields, outermost first. Requires `columnField`. */
	moreColumnFields?: string[];
	pageFilter?: PivotPageFilter;
	/** Hidden items per field (slicers, filter lists); the cache keeps every record. */
	itemFilters?: PivotItemFilter[];
	/** Number groups: the field shows buckets (`0-9`, `10-19`) instead of its values. */
	fieldGroups?: PivotFieldGroup[];
	/**
	 * Date range from a timeline. The field is one of `dateFields`.
	 */
	dateFilter?: PivotDateFilter;
	/**
	 * Cache fields that hold dates and sit on no axis (they exist for timelines): written as
	 * date items. Other date-valued fields stay plain numbers.
	 */
	dateFields?: string[];
	/** The workbook uses the 1904 date system (date groups need it to read serials). */
	date1904?: boolean;
	/** `x14:pivotCacheDefinition/@pivotCacheId`, the identity slicers use to find the cache. */
	slicerCacheId?: number;
	/**
	 * Captions in the language of the file or of the editor that authored the report. Omitted
	 * means English; a loaded file holds what its cached cells and attributes use.
	 */
	captions?: PivotCaptionOverrides;
	/** Item display names the file carries; they replace the item's value in the cells. */
	itemNames?: PivotItemName[];
	valueField: string;
	aggregate: PivotAggregate;
	valueCaption: string;
	/** Extra measures in Values axis order. The existing value fields remain the primary measure. */
	additionalValues?: PivotValueField[];
	/** Values on rows require multiple measures; any number of row and column fields. */
	valuesAxis?: 'columns' | 'rows';
	/**
	 * Where the row Values sit among the row fields: after them by default, `'first'` before
	 * them, `'between'` after the first one, or the number of row fields that precede them.
	 */
	valuesPosition?: PivotValuesPosition;
	/** Cached source fields and records. Refresh replaces these together with the output cells. */
	fields: string[];
	records: CellValue[][];
	location: CellRange;
	/** Loaded definitions remain verbatim until an explicit edit or refresh. */
	partName?: string;
	cacheId?: number;
	sourceXml?: string;
	dirty?: boolean;
}

export interface PivotTableOptions {
	source: string;
	rowField: string;
	layout?: PivotLayout;
	subtotalPosition?: 'top' | 'bottom';
	secondRowField?: string;
	moreRowFields?: string[];
	subtotals?: boolean;
	fieldSubtotals?: PivotFieldSubtotals[];
	dateGroups?: PivotDateGroup[];
	repeatLabels?: boolean;
	columnField?: string;
	moreColumnFields?: string[];
	pageFilter?: PivotPageFilter;
	fieldGroups?: PivotFieldGroup[];
	valueField: string;
	additionalValues?: PivotValueFieldOptions[];
	valuesAxis?: 'columns' | 'rows';
	valuesPosition?: PivotValuesPosition;
	aggregate?: PivotAggregate;
	name?: string;
	/** The captions of the new report; English when omitted. */
	captions?: PivotCaptionOverrides;
	/** The editor language (`de`, `es-ES`): the captions Excel writes in that UI language. */
	locale?: string;
}

export type PivotTablePatch = Partial<
	Pick<
		PivotTable,
		| 'rowField'
		| 'layout'
		| 'subtotalPosition'
		| 'secondRowField'
		| 'subtotals'
		| 'repeatLabels'
		| 'columnField'
		| 'valueField'
		| 'aggregate'
		| 'valuesAxis'
		| 'valuesPosition'
	>
> & {
	/** Replaces the third and following row fields (an empty list removes them). */
	moreRowFields?: string[];
	/** Replaces the second and following column fields (an empty list removes them). */
	moreColumnFields?: string[];
	/** Replaces every per-field subtotal setting (an empty list restores Automatic). */
	fieldSubtotals?: PivotFieldSubtotals[];
	/** Replaces every date group (an empty list ungroups the dates). */
	dateGroups?: PivotDateGroup[];
	additionalValues?: PivotValueFieldOptions[];
	pageFilter?: PivotPageFilter | null;
	/** Replaces every number group (an empty list ungroups). */
	fieldGroups?: PivotFieldGroup[];
};
