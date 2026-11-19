import { createCalcEngine } from './deps';
import type { Workbook } from '../model';
import type { EditContext } from './context';

export type CalcMode = 'auto' | 'manual';

/** The workbook's calculation mode (automatic unless the model says manual). */
export const calcModeOf = (workbook: Workbook): CalcMode =>
	workbook.calcMode === 'manual' ? 'manual' : 'auto';

/**
 * Runs the full calculation a workbook asks for when it opens (`calcPr fullCalcOnLoad`), unless it
 * calculates manually; true when it ran. Editors call it once on a workbook they open, before
 * showing it, in a worker or on the page.
 */
export function calculateOnOpen(workbook: Workbook): boolean {
	if (!workbook.fullCalcOnLoad || calcModeOf(workbook) !== 'auto') return false;
	createCalcEngine(workbook).recalculateAll();
	return true;
}

/**
 * Switches automatic or manual calculation (saved as `calcPr calcMode`). Switching back to
 * automatic recalculates, as Excel does. One undo step.
 */
export function setCalcMode(ctx: EditContext, mode: CalcMode): void {
	const { workbook } = ctx;
	if (calcModeOf(workbook) === mode) return;
	ctx.run(
		mode === 'manual' ? 'Manual calculation' : 'Automatic calculation',
		'view',
		[{ kind: 'meta' }],
		() => {
			if (mode === 'manual') workbook.calcMode = 'manual';
			else delete workbook.calcMode;
		},
	);
}
