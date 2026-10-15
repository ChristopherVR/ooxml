import type { CellRange } from './address';
import type { CellValue } from './model';

export type PivotAggregate = 'sum' | 'count' | 'average' | 'min' | 'max';
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

/** A worksheet-range pivot with up to two row fields and sixteen measures. */
export interface PivotTable {
	name: string;
	/** Qualified A1 source range, rewritten with the workbook's other references. */
	source: string;
	rowField: string;
	layout?: PivotLayout;
	subtotalPosition?: 'top' | 'bottom';
	secondRowField?: string;
	/** Automatic outer-field totals below each group. Defaults to true. */
	subtotals?: boolean;
	/** Repeat outer labels on detail rows. Defaults to false. */
	repeatLabels?: boolean;
	columnField?: string;
	pageFilter?: PivotPageFilter;
	/** Hidden items per field (slicers, filter lists); the cache keeps every record. */
	itemFilters?: PivotItemFilter[];
	/** `x14:pivotCacheDefinition/@pivotCacheId`, the identity slicers use to find the cache. */
	slicerCacheId?: number;
	valueField: string;
	aggregate: PivotAggregate;
	valueCaption: string;
	/** Extra measures in Values axis order. The existing value fields remain the primary measure. */
	additionalValues?: PivotValueField[];
	/** Values on rows require multiple measures and one or two real row fields. */
	valuesAxis?: 'columns' | 'rows';
	/** Row Values follow both real fields by default, or sit between/before two real fields. */
	valuesPosition?: 'last' | 'between' | 'first';
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
	subtotals?: boolean;
	repeatLabels?: boolean;
	columnField?: string;
	pageFilter?: PivotPageFilter;
	valueField: string;
	additionalValues?: PivotValueFieldOptions[];
	valuesAxis?: 'columns' | 'rows';
	valuesPosition?: 'last' | 'between' | 'first';
	aggregate?: PivotAggregate;
	name?: string;
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
> & { additionalValues?: PivotValueFieldOptions[]; pageFilter?: PivotPageFilter | null };
