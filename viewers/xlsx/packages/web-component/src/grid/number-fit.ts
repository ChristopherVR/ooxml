// A number's display fitted to its cell, as Excel does: General drops decimals or switches to
// scientific notation, any other format becomes `###`. The core's `formatValue` does the fitting
// with the browser measurer injected; `cellView` has no width input yet, so the grid asks here
// only for numbers whose 11-character General text overflows.
import {
	effectiveStyleId,
	formatValue,
	getCell,
	styleAt,
	type Workbook,
	type Worksheet,
} from '@christophervr/xlsx-core';

export function fitNumber(
	workbook: Workbook | undefined,
	sheet: Worksheet,
	row: number,
	col: number,
	width: number,
	measure: (text: string) => number,
): string | undefined {
	const value = getCell(sheet, row, col)?.value;
	if (typeof value !== 'number' || !workbook) return undefined;
	const numFmt = styleAt(workbook, effectiveStyleId(sheet, row, col)).numFmt || 'General';
	return formatValue(value, numFmt, { date1904: workbook.date1904, width, measure }).text;
}
