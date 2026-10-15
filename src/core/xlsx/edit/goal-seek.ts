// Goal Seek (Data > What-If Analysis): changes one input cell until a formula cell reaches a
// value, by iterating the recalculation engine. One undo step; the caller keeps the result
// (OK) or undoes it (Cancel).
import type { CellAddress } from '../address';
import { getCell } from '../cells';
import type { CellValue, Workbook } from '../model';
import { type EditContext, ensureCell, sheetAt } from './context';
import {
	DEFAULT_GOAL_ITERATIONS,
	DEFAULT_GOAL_MAX_CHANGE,
	type GoalSolverOptions,
	solveGoal,
} from './goal-solver';

export interface GoalSeekRequest {
	sheet: number;
	/** The "Set cell": must hold a formula. */
	target: CellAddress;
	/** The "To value". */
	goal: number;
	/** The "By changing cell": must hold a constant (or be empty). */
	changing: CellAddress;
}

/** Why a request cannot run, as Excel words it. */
export type GoalSeekProblem =
	| 'target-not-formula'
	| 'changing-has-formula'
	| 'same-cell'
	| 'goal-not-number';

export interface GoalSeekResult {
	/** Whether the target came within the maximum change of the goal. */
	found: boolean;
	/** The value left in the changing cell. */
	changingValue: number;
	/** The target's value after the search. */
	currentValue: CellValue;
	iterations: number;
}

/** The iteration settings Goal Seek uses: the workbook's, else Excel's defaults. */
export function goalSeekSettings(workbook: Workbook): Required<GoalSolverOptions> {
	return {
		maxIterations: workbook.iterate?.count ?? DEFAULT_GOAL_ITERATIONS,
		maxChange: workbook.iterate?.delta ?? DEFAULT_GOAL_MAX_CHANGE,
	};
}

export function checkGoalSeek(
	workbook: Workbook,
	request: GoalSeekRequest,
): GoalSeekProblem | null {
	const sheet = sheetAt(workbook, request.sheet);
	if (!Number.isFinite(request.goal)) return 'goal-not-number';
	const { target, changing } = request;
	if (target.row === changing.row && target.col === changing.col) return 'same-cell';
	if (!getCell(sheet, target.row, target.col)?.formula) return 'target-not-formula';
	if (getCell(sheet, changing.row, changing.col)?.formula) return 'changing-has-formula';
	return null;
}

const MESSAGES: Record<GoalSeekProblem, string> = {
	'target-not-formula': 'The Set cell must contain a formula.',
	'changing-has-formula': 'The By changing cell must contain a value, not a formula.',
	'same-cell': 'The Set cell and the By changing cell must be different cells.',
	'goal-not-number': 'The To value must be a number.',
};

export const goalSeekMessage = (problem: GoalSeekProblem): string => MESSAGES[problem];

export function goalSeek(ctx: EditContext, request: GoalSeekRequest): GoalSeekResult {
	const problem = checkGoalSeek(ctx.workbook, request);
	if (problem) throw new Error(MESSAGES[problem]);
	const { sheet: s, target, changing, goal } = request;
	const sheet = sheetAt(ctx.workbook, s);
	const settings = goalSeekSettings(ctx.workbook);
	const at = { sheet: s, row: changing.row, col: changing.col };
	const range = { start: changing, end: changing };
	return ctx.run(
		'Goal Seek',
		'cells',
		[{ kind: 'cells', sheet: s, ranges: [range] }],
		() => {
			const cell = ensureCell(sheet, changing.row, changing.col);
			const startValue = typeof cell.value === 'number' ? cell.value : 0;
			const apply = (x: number): void => {
				cell.value = x;
				ctx.calc.recalculateFrom([at]);
			};
			const read = (): CellValue => getCell(sheet, target.row, target.col)?.value ?? null;
			const solution = solveGoal(
				(x) => {
					apply(x);
					const value = read();
					return typeof value === 'number' ? value : Number.NaN;
				},
				goal,
				startValue,
				settings,
			);
			apply(solution.x);
			return {
				found: solution.found,
				changingValue: solution.x,
				currentValue: read(),
				iterations: solution.iterations,
			};
		},
		{ sheet: s, ranges: [range] },
	);
}
