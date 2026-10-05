import type { CellRange } from '../address.js';
import type { Workbook } from '../model.js';
import type { FormulaAst, FormulaError } from './ast.js';
import type { Matrix, RefValue, Scalar, Scope, Value } from './values.js';

/** What the evaluator needs from its environment; the calc engine implements it. */
export interface EvalHost {
	readonly workbook: Workbook;
	/** A cell's value, computing a pending formula first. */
	readCell(sheet: number, row: number, col: number): Scalar;
	/** Visits stored cells with a non-null value inside a range, in row-major order. */
	forEachStored(
		sheet: number,
		range: CellRange,
		visit: (value: Scalar, row: number, col: number) => void,
	): void;
	/**
	 * The values of a range as a matrix, cached for the rest of a recalculation (criteria ranges
	 * that many formulas read). Optional: callers fall back to `readCell`.
	 */
	readBlock?(sheet: number, range: CellRange): Matrix;
	/** One past the last stored row and column of a sheet. */
	bounds(sheet: number): { rows: number; cols: number };
	/** The spill range anchored at a cell, if it currently spills. */
	spillRange(sheet: number, row: number, col: number): CellRange | undefined;
	/** The formula text of a cell, if it has one. */
	cellFormula(sheet: number, row: number, col: number): string | undefined;
	/** Parses with caching. */
	parse(formula: string): FormulaAst | FormulaError;
	now(): Date;
	random(): number;
}

/** One evaluation: where the formula lives and which LET / LAMBDA variables are in scope. */
export interface Frame {
	readonly host: EvalHost;
	readonly sheet: number;
	readonly row: number;
	readonly col: number;
	readonly scope: Scope | undefined;
	/** Defined-name nesting depth (guards against names that refer to themselves). */
	readonly depth: number;
	/** The root node, for Excel's final-result rounding of additions and subtractions. */
	readonly root?: FormulaAst | undefined;
	/**
	 * Pre-dynamic-array (legacy) evaluation: a multi-cell reference used where one value is
	 * expected (an operator operand, a `value` parameter, the final result) is reduced by implicit
	 * intersection. Arguments of `any` (array) parameters are evaluated without it, as in Excel.
	 */
	readonly legacy?: boolean;
}

/** A lazily evaluated argument (IF, IFERROR, CHOOSE, LET, LAMBDA, ...). */
export interface LazyArg {
	readonly node: FormulaAst;
	get(): Value;
}

/** What function implementations receive besides their arguments. */
export interface CallContext {
	readonly frame: Frame;
	readonly workbook: Workbook;
	readonly sheet: number;
	readonly row: number;
	readonly col: number;
	readonly date1904: boolean;
	/** A dense matrix of a value (refs are read cell by cell). Throws `#VALUE!` for unions. */
	toMatrix(value: Value): Matrix;
	/** A single value by implicit intersection. */
	toScalar(value: Value): Scalar;
	/**
	 * Visits the values of an argument: refs sparsely (stored cells only), arrays fully, and
	 * a scalar argument once with `kind` `direct`.
	 */
	forEach(value: Value, visit: (value: Scalar, kind: 'direct' | 'ref' | 'array') => void): void;
	readCell(sheet: number, row: number, col: number): Scalar;
	/** Evaluates a node in this call's frame, optionally with extra variables. */
	evaluate(node: FormulaAst, scope?: Scope): Value;
	/** Calls a LAMBDA value. */
	callLambda(fn: Value, args: Value[]): Value;
	/** Resolves reference text (`A1`, `Sheet2!B2:C3`, a defined name) for INDIRECT. */
	parseReference(text: string): RefValue | undefined;
}
