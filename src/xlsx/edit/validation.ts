import { parseRange, rangeContains } from '../address.js';
import { forEachCellInRange } from '../cells.js';
import type { CellValue, ConditionalOperator, DataValidation, Workbook } from '../model.js';
import { isCellError } from '../model.js';
import { sheetByName } from '../workbook.js';
import { displayText, rawText } from './context.js';
import { type CalcEngine, createCalcEngine } from './deps.js';
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

/** Splits a list source `"a,b,c"` or resolves a range or name to its non-blank display texts. */
function listSource(
	workbook: Workbook,
	sheet: number,
	formula: string,
	evaluate: NonNullable<ValidationOptions['evaluate']> | undefined,
	at: { sheet: number; row: number; col: number },
	depth = 0,
): string[] | undefined {
	const text = formula.replace(/^=/, '').trim();
	const literal = /^"((?:[^"]|"")*)"$/.exec(text);
	if (literal)
		return (literal[1] ?? '')
			.replace(/""/g, '"')
			.split(',')
			.map((s) => s.trim())
			.filter((s) => s !== '');
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
		const out: string[] = [];
		forEachCellInRange(target, range, (cell) => {
			const shown = displayText(workbook, cell);
			if (shown !== '') out.push(shown);
		});
		return out;
	}
	if (!evaluate) return undefined;
	try {
		const value = evaluate(text, at);
		return value === null ? [] : [rawText(value)];
	} catch {
		return undefined;
	}
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
	return listSource(workbook, sheet, dv.formula1, options.evaluate ?? defaultEvaluate(workbook), {
		sheet,
		row,
		col,
	});
}

function check(
	workbook: Workbook,
	dv: DataValidation,
	value: CellValue,
	at: { sheet: number; row: number; col: number },
	evaluate: NonNullable<ValidationOptions['evaluate']>,
): boolean {
	if (dv.type === 'list') {
		const items =
			dv.formula1 === undefined
				? []
				: (listSource(workbook, at.sheet, dv.formula1, evaluate, at) ?? []);
		const text = rawText(value).toLowerCase();
		return items.some((item) => item.toLowerCase() === text);
	}
	if (dv.type === 'custom') {
		const result = operand(dv.formula1, evaluate, at);
		return result === true || (typeof result === 'number' && result !== 0);
	}
	const a = asNumber(operand(dv.formula1, evaluate, at) ?? null);
	const b = asNumber(operand(dv.formula2, evaluate, at) ?? null);
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
 * is explicitly false; rules whose error alert is turned off never reject.
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
	if (!dv || dv.type === 'none' || dv.showErrorMessage === false) return { ok: true };
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
