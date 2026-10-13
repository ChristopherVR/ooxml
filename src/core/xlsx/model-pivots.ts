import type { CellRange } from './address';
import type { CellValue } from './model';

export type PivotAggregate = 'sum' | 'count' | 'average' | 'min' | 'max';
export type PivotLayout = 'tabular' | 'compact' | 'outline';

/** A worksheet-range pivot with up to two row fields and one value field. */
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
	>
>;
