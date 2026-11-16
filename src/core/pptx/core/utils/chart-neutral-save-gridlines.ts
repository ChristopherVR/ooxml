/**
 * The chart-wide gridline flag (`style.hasGridlines`) saved through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C3). The flag is read off the first value axis's
 * `c:majorGridlines`; the object-tree writers never wrote it, so a flag edited on its own was lost
 * on save. It now sets the major gridlines of that axis (`setChartAxis`), as the shared
 * inspector's Show Gridlines patch does; a flag that only mirrors an axis edit is claimed with it.
 *
 * @module utils/chart-neutral-save-gridlines
 */

import { setChartAxis } from '../../../chart/index';
import type { PptxChartData } from '../types';
import { revertStyle } from './chart-neutral-save-common';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

export function planNeutralGridlineFlag(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	const was = before.style?.hasGridlines ?? false;
	const now = probe.style?.hasGridlines ?? false;
	if (was === now) {
		// `false` and absent read the same: keep the authored form.
		if (before.style?.hasGridlines !== probe.style?.hasGridlines)
			revertStyle(probe, before, ['hasGridlines']);
		return;
	}
	const axis = probe.axes?.find((item) => item.axisType === 'valAx');
	if (axis?.axisId === undefined) return;
	// An axis edit the axes planner claimed already writes the same gridlines.
	if ((axis.majorGridlines ?? false) !== now) {
		const id = axis.axisId;
		edits.push((space) => setChartAxis(space, { id }, { majorGridlines: now }));
	}
	revertStyle(probe, before, ['hasGridlines']);
}
