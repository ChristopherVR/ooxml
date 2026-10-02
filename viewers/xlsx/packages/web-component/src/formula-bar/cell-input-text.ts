// The text the formula bar (and the in-cell editor on F2) shows for a cell: what the user would
// type to recreate it. Formatting goes through the core number formatter and input parser.
import {
	effectiveStyleId,
	formatValue,
	getCell,
	isCellError,
	isDateFormat,
	parseCellInput,
	styleAt,
	type Workbook,
} from '@christophervr/xlsx-core';

/** Excel shows 15 significant digits in the formula bar. */
const plainNumber = (value: number): string => String(Number(value.toPrecision(15)));

function numberText(value: number, numFmt: string, date1904: boolean): string {
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

/** The raw, re-typeable content of a cell (formulas with their '='). */
export function cellInputText(workbook: Workbook, sheet: number, row: number, col: number): string {
	const ws = workbook.sheets[sheet];
	if (!ws) return '';
	const cell = getCell(ws, row, col);
	if (!cell) return '';
	if (cell.formula !== undefined) return `=${cell.formula}`;
	const value = cell.value;
	if (value === null) return '';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (isCellError(value)) return value.error;
	if (typeof value === 'number') {
		const style = styleAt(workbook, effectiveStyleId(ws, row, col));
		return numberText(value, style.numFmt || 'General', workbook.date1904);
	}
	const text = cell.richText?.length ? cell.richText.map((run) => run.text).join('') : value;
	if (text === '') return '';
	const parsed = parseCellInput(text, { date1904: workbook.date1904 });
	const literal = parsed.formula === undefined && parsed.value === text;
	return literal ? text : `'${text}`;
}
