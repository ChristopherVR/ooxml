// What bounds a whole recalculation, as opposed to one formula (see `withFormulaBudget`): the time
// it may run and the results it may store. Either stops it between formulas, leaving the formulas
// it did not reach with their previous values, rather than turning later cells into errors.
import type { CalculationStop } from './engine-types';
import { Matrix, type Scalar } from './values';

/**
 * The most array cells and text characters one recalculation may store as results, so formulas
 * that each stay within their own budget still cannot together exhaust memory. Both are far above
 * any workbook Excel opens comfortably; past them Excel stops with "out of resources" too.
 */
export const MAX_STORED_CELLS = 50_000_000;
export const MAX_STORED_TEXT = 536_870_912;

/** How many formulas run between two looks at the clock. */
const CLOCK_EVERY = 32;

/** Thrown between formulas to end a recalculation; the engine catches it and records why. */
export class CalculationStopped {
	constructor(readonly reason: CalculationStop) {}
}

/** The text a result holds (a padded matrix repeats its fill). */
function textOf(result: Scalar | Matrix): number {
	if (typeof result === 'string') return result.length;
	if (!(result instanceof Matrix)) return 0;
	let chars = 0;
	for (let r = 0; r < result.blockRows; r++)
		for (let c = 0; c < result.blockCols; c++) {
			const value = result.get(r, c);
			if (typeof value === 'string') chars += value.length;
		}
	const fill = result.fill;
	if (typeof fill === 'string')
		chars += fill.length * (result.rows * result.cols - result.blockRows * result.blockCols);
	return chars;
}

export class RecalcLimits {
	private deadline: number | undefined;
	private untilClock = CLOCK_EVERY;
	private cells = 0;
	private text = 0;

	constructor(private readonly timeLimitMs: number | undefined) {}

	/** Starts the limits of a new recalculation. */
	begin(): void {
		this.deadline =
			this.timeLimitMs === undefined ? undefined : performance.now() + this.timeLimitMs;
		this.untilClock = CLOCK_EVERY;
		this.cells = 0;
		this.text = 0;
	}

	/** Called before each top-level formula: stops the recalculation once its time is up. */
	tick(): void {
		if (this.deadline === undefined || --this.untilClock > 0) return;
		this.untilClock = CLOCK_EVERY;
		if (performance.now() > this.deadline) throw new CalculationStopped('time');
	}

	/** Counts a result about to be stored (`cells` it fills); stops past the stored limits. */
	store(result: Scalar | Matrix, cells: number): void {
		this.cells += cells;
		this.text += textOf(result);
		if (this.cells > MAX_STORED_CELLS || this.text > MAX_STORED_TEXT)
			throw new CalculationStopped('resources');
	}
}
