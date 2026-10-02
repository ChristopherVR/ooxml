import type { ErrorCode } from '../model.js';

/** Thrown by {@link parseFormula} (and the tokenizer) for malformed formulas. */
export class FormulaError extends Error {
	constructor(
		message: string,
		/** Character offset in the formula where the problem was found. */
		readonly position = -1,
	) {
		super(message);
		this.name = 'FormulaError';
	}
}

/** One end of a reference. Whole-column refs have no row, whole-row refs no column. */
export interface RefCorner {
	row: number;
	col: number;
	rowAbs: boolean;
	colAbs: boolean;
}

/** A parsed A1-style reference without its sheet prefix. */
export interface RefSpec {
	kind: 'cell' | 'area' | 'cols' | 'rows';
	start: RefCorner;
	/** Same as `start` for a single cell. */
	end: RefCorner;
}

/** The sheet part of a reference: `Sheet1!`, `'My Sheet'!`, `Sheet1:Sheet3!`, `[1]Sheet1!`. */
export interface SheetPrefix {
	/** Raw prefix text including the trailing `!`, preserved when formulas are rewritten. */
	text: string;
	sheet: string;
	/** Last sheet of a 3D reference. */
	sheet2?: string;
	/** External workbook index or name (`[1]`). */
	book?: string;
}

export type StructuredSpecial = '#All' | '#Data' | '#Headers' | '#Totals' | '#This Row';

/** A structured (table) reference such as `Table1[[#This Row],[Amount]]`. */
export interface StructuredRef {
	table?: string;
	specials: StructuredSpecial[];
	/** First (and last, for `[A]:[B]`) column name; absent means every column. */
	column?: string;
	column2?: string;
}

export type BinaryOperator =
	| '+'
	| '-'
	| '*'
	| '/'
	| '^'
	| '&'
	| '='
	| '<>'
	| '<'
	| '>'
	| '<='
	| '>='
	| ':'
	| ' '
	| ',';

/** The formula syntax tree produced by {@link parseFormula}. */
export type FormulaAst =
	| { type: 'number'; value: number }
	| { type: 'string'; value: string }
	| { type: 'boolean'; value: boolean }
	| { type: 'error'; code: ErrorCode }
	| {
			type: 'ref';
			prefix?: SheetPrefix;
			/** Absent for `#REF!` (optionally sheet-qualified). */
			ref?: RefSpec;
			/** `A1#` spill reference. */
			spill?: boolean;
	  }
	| { type: 'name'; prefix?: SheetPrefix; name: string }
	| { type: 'structured'; ref: StructuredRef; text: string }
	| { type: 'array'; rows: ArrayConstant[][] }
	| { type: 'unary'; op: '-' | '+' | '@'; operand: FormulaAst }
	| { type: 'percent'; operand: FormulaAst }
	| { type: 'binary'; op: BinaryOperator; left: FormulaAst; right: FormulaAst }
	| {
			type: 'call';
			/** Upper case, without `_xlfn.` / `_xlws.` prefixes. */
			name: string;
			/** As written. */
			rawName: string;
			args: FormulaAst[];
	  }
	| { type: 'invoke'; callee: FormulaAst; args: FormulaAst[] }
	| { type: 'missing' };

export type ArrayConstant = number | string | boolean | { error: ErrorCode };

/** Strips the storage prefixes Excel writes for newer functions (`_xlfn.`, `_xlws.`, `_xludf.`). */
export function normalizeFunctionName(name: string): string {
	return name.replace(/^(?:_xlfn\.|_xlws\.|_xludf\.)+/i, '').toUpperCase();
}

/** Strips the `_xlpm.` prefix Excel writes for LET and LAMBDA parameter names. */
export function normalizeLocalName(name: string): string {
	return name.replace(/^_xlpm\./i, '').toUpperCase();
}
