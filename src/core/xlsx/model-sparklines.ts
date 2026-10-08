// Excel 2010 sparklines (`x14:sparklineGroups` in a worksheet's extension list). The extension XML
// stays in `Worksheet.preserved` and is saved verbatim; this model is read from it on demand (see
// `sheetSparklineGroups`), so row, column and sheet edits that rewrite the preserved XML are always
// reflected and an untouched workbook saves exactly as before.
import type { CellRange } from './address';
import type { Color } from './model';

/** `line`, `column`, or `stacked` (Excel's win/loss). */
export type SparklineType = 'line' | 'column' | 'stacked';

/** How empty cells in the data range are shown: as a gap, as zero, or bridged by the line. */
export type SparklineEmptyCells = 'gap' | 'zero' | 'span';

/** Each sparkline scales on its own data, every sparkline of the group shares one scale, or fixed. */
export type SparklineAxisType = 'individual' | 'group' | 'custom';

/** One sparkline: its data range formula and the cell it is drawn in. */
export interface Sparkline {
	/** The data range formula (`xm:f`), for example `Sheet1!A1:E1`; absent for an empty sparkline. */
	formula?: string;
	/** The host cell (`xm:sqref`). */
	host: CellRange;
}

/** One `x14:sparklineGroup`: the settings its sparklines share. Absent attributes take defaults. */
export interface SparklineGroup {
	/** Default `line`. */
	type: SparklineType;
	/** Line width in points (default 0.75). */
	lineWeight: number;
	/** Plot points by the date range (`xm:f` of the group) instead of evenly. */
	dateAxis: boolean;
	/** The date range formula of a date axis. */
	dateFormula?: string;
	markers: boolean;
	high: boolean;
	low: boolean;
	first: boolean;
	last: boolean;
	negative: boolean;
	/** Default `zero` (the schema default; Excel writes `gap` explicitly). */
	displayEmptyCellsAs: SparklineEmptyCells;
	displayXAxis: boolean;
	/** Values in hidden rows and columns are plotted. */
	displayHidden: boolean;
	rightToLeft: boolean;
	minAxisType: SparklineAxisType;
	maxAxisType: SparklineAxisType;
	manualMin?: number;
	manualMax?: number;
	colorSeries?: Color;
	colorNegative?: Color;
	colorAxis?: Color;
	colorMarkers?: Color;
	colorFirst?: Color;
	colorLast?: Color;
	colorHigh?: Color;
	colorLow?: Color;
	sparklines: Sparkline[];
}
