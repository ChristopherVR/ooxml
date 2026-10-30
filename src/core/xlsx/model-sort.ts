// The saved sort of a sheet or auto-filter (`sortState`, ECMA-376 18.3.1.92).
import type { CellRange } from './address';
import type { DifferentialStyle } from './model';

export type SortBy = 'value' | 'cellColor' | 'fontColor' | 'icon';

/** One `sortCondition`: a sort level over one column (or one row when sorting left to right). */
export interface SortCondition {
	/** The column (or row) of the data this level orders by. */
	ref: CellRange;
	descending?: boolean;
	/** Absent means `value`. */
	sortBy?: SortBy;
	/** Name of a custom list, or its items joined by commas (Excel writes the latter). */
	customList?: string;
	dxfId?: number;
	/** The differential format `dxfId` named; colour sorts match its fill or font colour. */
	style?: DifferentialStyle;
	iconSet?: string;
	/** Zero-based icon within the set; absent means "No Cell Icon". */
	iconId?: number;
}

export interface SortState {
	range: CellRange;
	/** Sort left to right (columns are reordered). */
	columnSort?: boolean;
	caseSensitive?: boolean;
	conditions: SortCondition[];
}
