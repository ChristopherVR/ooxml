import type { Workbook } from '../model.js';
import type { EditContext } from './context.js';

export type CalcMode = 'auto' | 'manual';

/** The workbook's calculation mode (automatic unless the model says manual). */
export const calcModeOf = (workbook: Workbook): CalcMode =>
	workbook.calcMode === 'manual' ? 'manual' : 'auto';

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
