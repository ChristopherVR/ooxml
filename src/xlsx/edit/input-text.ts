// The text the formula bar (and the in-cell editor on F2) shows for a cell: what the user would
// type to recreate it. Moved from xlsx-viewer packages/web-component/src/formula-bar/
// cell-input-text.ts so find and replace work on the same text the formula bar shows.
import { getCell } from '../cells.js';
import type { Cell, Workbook } from '../model.js';
import { isCellError } from '../model.js';
import { styleAt } from '../styles.js';
import { formatValue, isDateFormat, parseCellInput } from './deps.js';

/** Excel shows 15 significant digits in the formula bar. */
const plainNumber = (value: number): string => String(Number(value.toPrecision(15)));

/** A number as the formula bar shows it: dates and times as US short input, percents with %. */
export function numberInputText(value: number, numFmt: string, date1904 = false): string {
	if (isDateFormat(numFmt)) {
		const whole = Math.floor(value);
		const hasTime = value !== whole;
		const hasDate = whole !== 0 || !hasTime;
		const format = hasDate ? (hasTime ? 'm/d/yyyy h:mm:ss AM/PM' : 'm/d/yyyy') : 'h:mm:ss AM/PM';
		return formatValue(value, format, { date1904 }).text;
	}
	if (/%/.test(numFmt.replace(/"[^"]*"|\\./g, ''))) return `${plainNumber(value * 100)}%`;
	return plainNumber(value);
}

/**
 * The formula-bar text of a cell. `quote` (the default) prefixes text that would otherwise be
 * read back as a number, date, boolean or formula with an apostrophe, as Excel's editor does.
 */
export function formulaBarText(
	workbook: Workbook,
	cell: Cell | undefined,
	options: { quote?: boolean } = {},
): string {
	if (!cell) return '';
	if (cell.formula !== undefined) return `=${cell.formula}`;
	const value = cell.value;
	if (value === null) return '';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (isCellError(value)) return value.error;
	if (typeof value === 'number')
		return numberInputText(
			value,
			styleAt(workbook, cell.styleId).numFmt || 'General',
			workbook.date1904,
		);
	const text = cell.richText?.length ? cell.richText.map((run) => run.text).join('') : value;
	if (text === '' || options.quote === false) return text;
	const parsed = parseCellInput(text, { date1904: workbook.date1904 });
	const literal = parsed.formula === undefined && parsed.value === text;
	return literal ? text : `'${text}`;
}

/** The re-typeable content of a cell (formulas with their '='), as the formula bar shows it. */
export function cellInputText(workbook: Workbook, sheet: number, row: number, col: number): string {
	const ws = workbook.sheets[sheet];
	return ws ? formulaBarText(workbook, getCell(ws, row, col)) : '';
}
