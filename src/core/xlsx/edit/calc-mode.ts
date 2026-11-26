import { type CalculationStop, createCalcEngine } from './deps';
import type { CalcExternalProviders } from '../formula/providers';
import type { Workbook } from '../model';
import type { EditContext } from './context';

export type CalcMode = 'auto' | 'manual';

/** The workbook's calculation mode (automatic unless the model says manual). */
export const calcModeOf = (workbook: Workbook): CalcMode =>
	workbook.calcMode === 'manual' ? 'manual' : 'auto';

/** Whether a workbook asks for a full calculation when it opens (`calcPr fullCalcOnLoad`). */
export const needsCalculationOnOpen = (workbook: Workbook): boolean =>
	!!workbook.fullCalcOnLoad && calcModeOf(workbook) === 'auto';

/** What `calculateOnOpen` did. */
export interface OpenCalculation {
	/** Whether the calculation ran: the workbook asks for it and calculates automatically. */
	ran: boolean;
	/**
	 * Why it stopped before the end, if it did: the formulas it did not reach keep the values the
	 * file was saved with, and the next recalculation computes everything.
	 */
	stopped?: CalculationStop;
}

/** How a calculation on open is bounded (see `CalcEngineOptions`). */
export interface OpenCalculationLimits {
	/** How long it may run. */
	timeLimitMs?: number;
	/** The most work one formula may need before the calculation stops. */
	workCap?: number;
	/** Host providers available to this calculation. */
	providers?: CalcExternalProviders;
}

/**
 * Runs the full calculation a workbook asks for when it opens (`calcPr fullCalcOnLoad`), unless it
 * calculates manually. Editors call it once on a workbook they open, before showing it, in a
 * worker or on the page, bounded by `limits`.
 */
export function calculateOnOpen(
	workbook: Workbook,
	limits: OpenCalculationLimits = {},
): OpenCalculation {
	if (!needsCalculationOnOpen(workbook)) return { ran: false };
	const engine = createCalcEngine(workbook, limits);
	engine.recalculateAll();
	return engine.stopped ? { ran: true, stopped: engine.stopped } : { ran: true };
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
