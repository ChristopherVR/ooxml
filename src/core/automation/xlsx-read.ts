// What the headless spreadsheet reads report besides raw values: each sheet's extent, visibility
// and tables, the workbook's defined names, and each cell's text as the grid displays it, so a
// caller can tell what to read and read dates and percentages without knowing Excel's encoding.
import { type CellRange, formatAddress, formatRange } from '../xlsx/address';
import { getCell, usedRange } from '../xlsx/cells';
import { displayText } from '../xlsx/display-text';
import type { Cell, CellValue, Workbook, Worksheet } from '../xlsx/model';
import { type NumberFormatView, numberFormatView } from '../xlsx/numfmt/view';
import { styleAt } from '../xlsx/styles';

/** The most tables per sheet, and defined names, that an inspection lists. */
export const INSPECT_LIST_LIMIT = 500;

/** Each sheet's used range (the cells it stores: values, formulas or formatting), state and tables. */
export function sheetSummaries(workbook: Workbook) {
	return workbook.sheets.map((sheet, index) => {
		const used = usedRange(sheet);
		const omitted = sheet.tables.length - INSPECT_LIST_LIMIT;
		return {
			index,
			name: sheet.name,
			state: sheet.state,
			usedRange: used ? formatRange(used) : null,
			rows: used ? used.end.row - used.start.row + 1 : 0,
			columns: used ? used.end.col - used.start.col + 1 : 0,
			tables: sheet.tables.slice(0, INSPECT_LIST_LIMIT).map((table) => ({
				name: table.displayName,
				range: formatRange(table.range),
				headerRow: table.headerRow,
				totalsRow: table.totalsRow,
			})),
			...(omitted > 0 ? { tablesOmitted: omitted } : {}),
		};
	});
}

/** The defined names Excel's Name Manager lists (hidden ones are left out), with their scope. */
export function definedNameList(workbook: Workbook) {
	const visible = workbook.definedNames.filter((name) => !name.hidden);
	const definedNames = visible.slice(0, INSPECT_LIST_LIMIT).map((name) => ({
		name: name.name,
		refersTo: name.formula,
		...(name.localSheet === undefined ? {} : { sheetIndex: name.localSheet }),
	}));
	const omitted = visible.length - INSPECT_LIST_LIMIT;
	return omitted > 0 ? { definedNames, definedNamesOmitted: omitted } : { definedNames };
}

export interface CellView extends NumberFormatView {
	value: CellValue;
	formula?: string;
	/** The text as the grid displays it, with the cell's number format applied. */
	text: string;
}

/** A cell's value and formula with its displayed text, number format and ISO date. */
export function cellView(workbook: Workbook, cell: Cell | undefined): CellView {
	const value = cell?.value ?? null;
	const text = displayText(workbook, cell);
	const view: CellView = cell?.formula ? { value, formula: cell.formula, text } : { value, text };
	const numFmt = styleAt(workbook, cell?.styleId).numFmt;
	return { ...view, ...numberFormatView(numFmt, value, workbook.date1904) };
}

/** The most cells one range or table read returns. */
const READ_LIMIT = 10_000;

/** The cells of `range`, row by row, each with its address and {@link cellView}. */
export function rangeCells(workbook: Workbook, sheet: Worksheet, range: CellRange) {
	const count = (range.end.row - range.start.row + 1) * (range.end.col - range.start.col + 1);
	if (count > READ_LIMIT) throw new Error(`Read at most ${READ_LIMIT} cells per request`);
	const cells: ({ address: string } & CellView)[] = [];
	for (let row = range.start.row; row <= range.end.row; row++)
		for (let col = range.start.col; col <= range.end.col; col++)
			cells.push({
				address: formatAddress({ row, col }),
				...cellView(workbook, getCell(sheet, row, col)),
			});
	return cells;
}
