import type { CellRange } from './address';
import type { CellValue } from './model';

export type PivotAggregate = 'sum' | 'count' | 'average' | 'min' | 'max';

/** One worksheet-range pivot with one row field, optional column field and one value field. */
export interface PivotTable {
	name: string;
	/** Qualified A1 source range, rewritten with the workbook's other references. */
	source: string;
	rowField: string;
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
	columnField?: string;
	valueField: string;
	aggregate?: PivotAggregate;
	name?: string;
}

export type PivotTablePatch = Partial<
	Pick<PivotTable, 'rowField' | 'columnField' | 'valueField' | 'aggregate'>
>;
