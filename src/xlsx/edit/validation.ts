import { rangeContains } from '../address.js';
import type { Cell, CellValue, ConditionalOperator, DataValidation, Workbook } from '../model.js';
import { isCellError } from '../model.js';
import { rawText } from './context.js';
import { type CalcEngine, createCalcEngine, translateFormula } from './deps.js';
import { listItemMatches, listSource } from './validation-list.js';
import type { ValidationResult } from './types.js';

export interface ValidationOptions {
	/** Evaluates a validation formula at the cell (the session passes its calc engine). */
	evaluate?: (formula: string, at: { sheet: number; row: number; col: number }) => CellValue;
}

export const DEFAULT_VALIDATION_MESSAGE =
	"This value doesn't match the data validation restrictions defined for this cell.";

const engines = new WeakMap<Workbook, CalcEngine>();
function defaultEvaluate(workbook: Workbook): NonNullable<ValidationOptions['evaluate']> {
	return (formula, at) => {
		let engine = engines.get(workbook);
		if (!engine) {
			engine = createCalcEngine(workbook);
			engines.set(workbook, engine);
		}
		return engine.evaluate(formula, at);
	};
}

/** The validation rule covering a cell, if any. */
export function validationAt(
	workbook: Workbook,
	sheet: number,
	row: number,
	col: number,
): DataValidation | undefined {
	return workbook.sheets[sheet]?.dataValidations.find((dv) =>
		dv.ranges.some((r) => rangeContains(r, { row, col })),
	);
}

/** A validation operand: a number or quoted literal directly, anything else evaluated. */
function operand(
	formula: string | undefined,
	evaluate: NonNullable<ValidationOptions['evaluate']>,
	at: { sheet: number; row: number; col: number },
): CellValue | undefined {
	if (formula === undefined) return undefined;
	const text = formula.replace(/^=/, '').trim();
	if (text !== '' && Number.isFinite(Number(text))) return Number(text);
	const quoted = /^"((?:[^"]|"")*)"$/.exec(text);
	if (quoted) return (quoted[1] ?? '').replace(/""/g, '"');
	try {
		return evaluate(text, at);
	} catch {
		return undefined;
	}
}

function compare(
	operator: ConditionalOperator,
	v: number,
	a: number,
	b: number | undefined,
): boolean {
	switch (operator) {
		case 'between':
			return v >= a && v <= (b ?? a);
		case 'notBetween':
			return v < a || v > (b ?? a);
		case 'equal':
			return v === a;
		case 'notEqual':
			return v !== a;
		case 'greaterThan':
			return v > a;
		case 'lessThan':
			return v < a;
		case 'greaterThanOrEqual':
			return v >= a;
		case 'lessThanOrEqual':
			return v <= a;
	}
}

function asNumber(value: CellValue): number | undefined {
	if (typeof value === 'number') return value;
	if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value)))
		return Number(value);
	return undefined;
}

/** The drop-down entries of a list validation on a cell, or undefined when it has none. */
export function listValidationOptions(
	workbook: Workbook,
	sheet: number,
	row: number,
	col: number,
	options: ValidationOptions = {},
): string[] | undefined {
	const dv = validationAt(workbook, sheet, row, col);
	if (!dv || dv.type !== 'list' || dv.showDropDown === false || dv.formula1 === undefined)
		return undefined;
	const at = { sheet, row, col };
	const items = listSource(
		workbook,
		sheet,
		relativeTo(dv, dv.formula1, at),
		options.evaluate ?? defaultEvaluate(workbook),
		at,
	);
	return items?.map((item) => item.text);
}

/**
 * A rule formula as it applies at a cell. SpreadsheetML stores validation formulas relative to
 * the top-left corner of the bounding box of the rule's ranges (Excel writes `sqref="C3:C5 A1:A2"`
 * with formulas relative to A1), so relative references move with the target cell.
 */
function relativeTo(dv: DataValidation, formula: string, at: { row: number; col: number }): string {
	if (dv.ranges.length === 0) return formula;
	let top = Infinity;
	let left = Infinity;
	for (const r of dv.ranges) {
		top = Math.min(top, r.start.row, r.end.row);
		left = Math.min(left, r.start.col, r.end.col);
	}
	const dRow = at.row - top;
	const dCol = at.col - left;
	if (dRow === 0 && dCol === 0) return formula;
	const text = formula.replace(/^=/, '');
	return translateFormula(text, dRow, dCol);
}

/**
 * Runs `fn` with the proposed value standing in the target cell (formula removed), so rules
 * such as `ISNUMBER(A1)` on A1 see what is being entered rather than what is stored. The cell is
 * restored exactly afterwards.
 */
function withProposedValue<T>(
	workbook: Workbook,
	at: { sheet: number; row: number; col: number },
	value: CellValue,
	fn: () => T,
): T {
	const sheet = workbook.sheets[at.sheet];
	if (!sheet) return fn();
	let row = sheet.rows.get(at.row);
	const createdRow = !row;
	if (!row) {
		row = new Map<number, Cell>();
		sheet.rows.set(at.row, row);
	}
	const previous = row.get(at.col);
	const proposed: Cell = { value };
	if (previous?.styleId !== undefined) proposed.styleId = previous.styleId;
	row.set(at.col, proposed);
	try {
		return fn();
	} finally {
		if (previous) row.set(at.col, previous);
		else row.delete(at.col);
		if (createdRow) sheet.rows.delete(at.row);
	}
}

function check(
	workbook: Workbook,
	dv: DataValidation,
	value: CellValue,
	at: { sheet: number; row: number; col: number },
	evaluate: NonNullable<ValidationOptions['evaluate']>,
): boolean {
	const f1 = dv.formula1 === undefined ? undefined : relativeTo(dv, dv.formula1, at);
	const f2 = dv.formula2 === undefined ? undefined : relativeTo(dv, dv.formula2, at);
	if (dv.type === 'list') {
		const items = f1 === undefined ? [] : (listSource(workbook, at.sheet, f1, evaluate, at) ?? []);
		return items.some((item) => listItemMatches(item, value));
	}
	if (dv.type === 'custom') {
		const result = withProposedValue(workbook, at, value, () => operand(f1, evaluate, at));
		return result === true || (typeof result === 'number' && result !== 0);
	}
	const a = asNumber(operand(f1, evaluate, at) ?? null);
	const b = asNumber(operand(f2, evaluate, at) ?? null);
	if (a === undefined) return true;
	const operator = dv.operator ?? 'between';
	if (dv.type === 'textLength') return compare(operator, rawText(value).length, a, b);
	const n = asNumber(value);
	if (n === undefined || isCellError(value)) return false;
	if (dv.type === 'whole' && !Number.isInteger(n)) return false;
	return compare(operator, n, a, b);
}

/**
 * Checks a value against the data validation of a cell. Blank values pass unless `allowBlank`
 * is explicitly false. Only rules whose error alert is on reject: `showErrorMessage` defaults to
 * false in SpreadsheetML, so a rule read without the attribute is not enforced (as in Excel).
 */
export function validateCellInput(
	workbook: Workbook,
	sheet: number,
	row: number,
	col: number,
	value: CellValue,
	options: ValidationOptions = {},
): ValidationResult {
	const dv = validationAt(workbook, sheet, row, col);
	if (!dv || dv.type === 'none' || dv.showErrorMessage !== true) return { ok: true };
	if (value === null || value === '') return dv.allowBlank === false ? failure(dv) : { ok: true };
	const evaluate = options.evaluate ?? defaultEvaluate(workbook);
	return check(workbook, dv, value, { sheet, row, col }, evaluate) ? { ok: true } : failure(dv);
}

function failure(dv: DataValidation): ValidationResult {
	const result: ValidationResult = {
		ok: false,
		style: dv.errorStyle ?? 'stop',
		message: dv.error || DEFAULT_VALIDATION_MESSAGE,
	};
	if (dv.errorTitle) result.title = dv.errorTitle;
	return result;
}
