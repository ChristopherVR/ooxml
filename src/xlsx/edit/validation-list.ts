import { parseRange } from '../address.js';
import { forEachCellInRange } from '../cells.js';
import type { CellValue, Workbook } from '../model.js';
import { isCellError } from '../model.js';
import { sheetByName } from '../workbook.js';
import { displayText, rawText } from './context.js';
import { parseCellInput } from './deps.js';

export type Evaluate = (
	formula: string,
	at: { sheet: number; row: number; col: number },
) => CellValue;

/** One entry of a list validation: what the drop-down shows and the value it stands for. */
export interface ListItem {
	text: string;
	value: CellValue;
}

/** Splits a list source `"a,b,c"` or resolves a range or name to its non-blank entries. */
export function listSource(
	workbook: Workbook,
	sheet: number,
	formula: string,
	evaluate: Evaluate | undefined,
	at: { sheet: number; row: number; col: number },
	depth = 0,
): ListItem[] | undefined {
	const text = formula.replace(/^=/, '').trim();
	const literal = /^"((?:[^"]|"")*)"$/.exec(text);
	if (literal)
		return (literal[1] ?? '')
			.replace(/""/g, '"')
			.split(',')
			.map((s) => s.trim())
			.filter((s) => s !== '')
			.map((s) => ({ text: s, value: parseCellInput(s, { date1904: workbook.date1904 }).value }));
	const named = workbook.definedNames.find(
		(n) =>
			n.name.toLowerCase() === text.toLowerCase() &&
			(n.localSheet === undefined || n.localSheet === sheet),
	);
	if (named && depth < 4)
		return listSource(workbook, sheet, named.formula, evaluate, at, depth + 1);
	const ref = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(text);
	const sheetName = ref ? (ref[1]?.replace(/''/g, "'") ?? ref[2]) : undefined;
	const target = sheetName ? sheetByName(workbook, sheetName) : workbook.sheets[sheet];
	const range = parseRange((ref ? (ref[3] ?? '') : text).replace(/\$/g, ''));
	if (target && range) {
		const out: ListItem[] = [];
		forEachCellInRange(target, range, (cell) => {
			const shown = displayText(workbook, cell);
			if (shown !== '') out.push({ text: shown, value: cell.value });
		});
		return out;
	}
	if (!evaluate) return undefined;
	try {
		const value = evaluate(text, at);
		return value === null ? [] : [{ text: rawText(value), value }];
	} catch {
		return undefined;
	}
}

/**
 * Whether a typed value matches a list entry the way Excel compares them: numbers (and dates,
 * which are serials) by value, text case-insensitively, booleans and errors by identity. A text
 * entry never matches a number, so `1.5` matches a cell holding 1.5 shown as `1.50`.
 */
export function listItemMatches(item: ListItem, value: CellValue): boolean {
	const target = item.value;
	if (typeof value === 'number')
		return typeof target === 'number' && Math.abs(target - value) <= 1e-12 * Math.abs(value);
	if (typeof value === 'boolean') return target === value;
	if (isCellError(value)) return isCellError(target) && target.error === value.error;
	if (typeof value === 'string') {
		if (typeof target !== 'string') return false;
		return target.toLowerCase() === value.toLowerCase();
	}
	return false;
}
