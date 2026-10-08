import type { CellRange } from '../address';
import type { CellValue } from '../model';

export interface CellPosition {
	sheet: number;
	row: number;
	col: number;
}

export interface CalcEngine {
	/** Evaluates every formula cell and sets `cell.value`. */
	recalculateAll(): void;
	/** Recalculates the formulas that depend on the changed cells (and volatile formulas). */
	recalculateFrom(changes: CellPosition[]): void;
	/** Evaluates an ad hoc formula at a position (status bar, conditional formats, validation). */
	evaluate(formula: string, at: CellPosition): CellValue;
	/** Forgets the dependency graph (after structural edits); the next recalculation rebuilds it. */
	invalidate(): void;
	/**
	 * Builds the dependency graph ahead of the first recalculation, in slices: work stops once
	 * `timeRemaining` (milliseconds left, like an idle deadline's) reaches zero and resumes on the
	 * next call. Returns true when the graph is ready (or there is nothing to prepare). Without
	 * `timeRemaining` it finishes. A recalculation that comes first completes the work itself.
	 */
	prepare(timeRemaining?: () => number): boolean;
	/**
	 * Drops a graph prepared before the first recalculation, for a workbook changed without the
	 * engine being told (edits in manual calculation mode); the next recalculation rebuilds it.
	 */
	discardPreparation(): void;
	/** Like `evaluate` but keeps array results (list validation sources, chart series). */
	evaluateArray(formula: string, at: CellPosition): CellValue[][];
	/** The current spill range anchored at a cell. */
	spillRange(sheet: number, row: number, col: number): CellRange | undefined;
	/** Formula cells found in circular references during the last recalculation. */
	circularCells(): CellPosition[];
}

export interface CalcEngineOptions {
	/** Clock for NOW and TODAY (tests). */
	now?: () => Date;
	/** Random source for RAND, RANDBETWEEN and RANDARRAY (tests). */
	random?: () => number;
}

export const CIRCULAR_REFERENCE_WARNING =
	'The workbook contains circular references; the cells involved show 0.';
