import type { CellRange } from './address';
import type { CellValue } from './model';

export type PivotAggregate = 'sum' | 'count' | 'average' | 'min' | 'max';
export type PivotLayout = 'tabular' | 'compact' | 'outline';
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
	valueField: string;
	aggregate: PivotAggregate;
	valueCaption: string;
	/** Extra measures in Values axis order. The existing value fields remain the primary measure. */
	additionalValues?: PivotValueField[];
	/** Values on rows currently require multiple measures, tabular form and one real row field. */
	valuesAxis?: 'columns' | 'rows';
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
	valueField: string;
	additionalValues?: PivotValueFieldOptions[];
	valuesAxis?: 'columns' | 'rows';
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
	>
> & { additionalValues?: PivotValueFieldOptions[] };
