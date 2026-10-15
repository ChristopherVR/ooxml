// What-If Analysis model: data tables (`<f t="dataTable">`) and scenarios (`<scenarios>`).
import type { CellRange } from './address';

/**
 * A what-if data table over the result range `ref` (the interior, without the input header
 * row and column). Excel shows it as the array formula `{=TABLE(row input, column input)}`.
 * - Two-variable (`dt2D`): `r1` is the row input cell and `r2` the column input cell; the
 *   formula sits in the corner cell above the left column and left of the top row.
 * - One-variable, values down the left column (`dtr` absent): `r1` is the column input cell
 *   and the formulas sit in the row above `ref`.
 * - One-variable, values across the top row (`dtr`): `r1` is the row input cell and the
 *   formulas sit in the column left of `ref`.
 * Input cells are on the same sheet as the table. `del1` and `del2` mark deleted input cells.
 */
export interface DataTable {
	ref: CellRange;
	/** A1 reference of the first input cell. */
	r1: string;
	/** A1 reference of the column input cell of a two-variable table. */
	r2?: string;
	dt2D?: true;
	dtr?: true;
	del1?: true;
	del2?: true;
}

/** One changing cell of a scenario with the value it takes (`inputCells`). */
export interface ScenarioCell {
	/** A1 reference of the changing cell. */
	ref: string;
	/** The value as text, exactly as stored (`val`). */
	value: string;
	deleted?: true;
	undone?: true;
	numFmtId?: number;
}

export interface Scenario {
	name: string;
	locked?: boolean;
	hidden?: boolean;
	/** Who last changed it. */
	user?: string;
	comment?: string;
	cells: ScenarioCell[];
}

/** The sheet's scenarios (`<scenarios>`). `current` and `show` are zero-based list indexes. */
export interface ScenarioSet {
	current?: number;
	show?: number;
	/** The result cells (`sqref`). */
	sqref?: string;
	list: Scenario[];
}
