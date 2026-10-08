// Format-neutral model of the chart chrome that sits beside the plot (ECMA-376 Part 1, 21.2): the
// floor and walls of a 3-D chart (`c:floor`, `c:sideWall`, `c:backWall`, CT_Surface), the data
// table under the plot (`c:dTable`, CT_DTable) and a value axis's display units (`c:dispUnits`,
// CT_DispUnits).
import type { ChartLayout, ChartShapeProperties, ChartText, ChartTextBody } from './model-series';

/** A floor or wall of a 3-D chart (`c:floor`, `c:sideWall`, `c:backWall`). */
export interface ChartSurface {
	/** `c:thickness/@val`, percent of the chart depth. */
	thickness?: number;
	spPr?: ChartShapeProperties;
	/** `c:pictureOptions` as written, kept for round-trip. */
	pictureOptionsXml?: string;
	extLst?: string;
}

/** The data table drawn under the plot (`c:dTable`). */
export interface ChartDataTable {
	showHorizontalBorder?: boolean;
	showVerticalBorder?: boolean;
	showOutline?: boolean;
	/** Whether the legend keys are drawn beside the series names. */
	showKeys?: boolean;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
	extLst?: string;
}

/** The label of a value axis's display units (`c:dispUnitsLbl`). */
export interface ChartDisplayUnitsLabel {
	layout?: ChartLayout;
	tx?: ChartText;
	spPr?: ChartShapeProperties;
	txPr?: ChartTextBody;
}

/** Display units of a value axis (`c:dispUnits`): one of `builtInUnit` and `customUnit`. */
export interface ChartDisplayUnits {
	/** `hundreds`, `thousands`, `millions`... (ST_BuiltInUnit), as written. */
	builtInUnit?: string;
	/** `c:custUnit/@val`, the divisor. */
	customUnit?: number;
	/** Present when the units label is shown. */
	label?: ChartDisplayUnitsLabel;
	extLst?: string;
}
