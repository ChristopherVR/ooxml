// What the headless spreadsheet reads report besides raw values: each sheet's extent, visibility
// and tables, the workbook's defined names, and each cell's text as the grid displays it, so a
// caller can tell what to read and read dates and percentages without knowing Excel's encoding.
import { type CellRange, formatAddress, formatRange } from '../xlsx/address';
import { getCell, usedRange } from '../xlsx/cells';
import { displayText } from '../xlsx/display-text';
import type { Cell, CellValue, Workbook, Worksheet } from '../xlsx/model';
import { MAX_DATE_SERIAL, serialToDate } from '../xlsx/numfmt/date';
import { dateTimeParts } from '../xlsx/numfmt/format';
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

/**
 * ISO 8601 without a time zone for a date or time serial: a date when it is whole and the format
 * shows no time, a time of day when there is no calendar day to name (below 1 in the 1900 system,
 * whose day 0 is fictitious, or a time-only format in the 1904 one), both otherwise, with
 * milliseconds when there are any. `undefined` before the epoch, after 9999 and for the fictitious
 * 1900-02-29 (serial 60).
 */
function isoOf(
	serial: number,
	date1904: boolean,
	shows: { date: boolean; time: boolean },
): string | undefined {
	// `!(serial >= 0)` also refuses NaN; past 9999 a date may not exist in JavaScript at all.
	if (!(serial >= 0) || serial > MAX_DATE_SERIAL + 1) return undefined;
	if (!date1904 && serial >= 60 && serial < 61) return undefined;
	const date = serialToDate(serial, date1904);
	if (date.getUTCFullYear() > 9999) return undefined;
	const iso = date.toISOString();
	const time = iso.slice(11, iso.endsWith('.000Z') ? 19 : 23);
	if (serial < 1 && (!date1904 || !shows.date)) return time;
	return Number.isInteger(serial) && !shows.time ? iso.slice(0, 10) : `${iso.slice(0, 10)}T${time}`;
}

export interface CellView {
	value: CellValue;
	formula?: string;
	/** The text as the grid displays it, with the cell's number format applied. */
	text: string;
	/** The number format code, absent for General. */
	numberFormat?: string;
	/**
	 * The value of a cell shown as a date or time, in ISO 8601 without a time zone; absent for
	 * numbers and durations (`[h]:mm`).
	 */
	iso?: string;
}

/** A cell's value and formula with its displayed text, number format and ISO date. */
export function cellView(workbook: Workbook, cell: Cell | undefined): CellView {
	const value = cell?.value ?? null;
	const text = displayText(workbook, cell);
	const view: CellView = cell?.formula ? { value, formula: cell.formula, text } : { value, text };
	const numFmt = styleAt(workbook, cell?.styleId).numFmt;
	if (numFmt === 'General') return view;
	view.numberFormat = numFmt;
	const parts = typeof value === 'number' ? dateTimeParts(numFmt, value) : undefined;
	if (parts && typeof value === 'number') {
		const iso = isoOf(value, workbook.date1904, parts);
		if (iso) view.iso = iso;
	}
	return view;
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
