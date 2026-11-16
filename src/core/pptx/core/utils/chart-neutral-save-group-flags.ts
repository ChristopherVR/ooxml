/**
 * Chart group settings edited on their own, saved through the shared chart core
 * (docs/agnostic-core-plan.md, chart step C3): `varyColors`, the bar direction, the scatter style
 * and the `standard` grouping flag of a kept chart type. The object-tree writers wrote none of the
 * first three on an existing part (the edit was lost on save) and wrote the grouping flag only
 * where the container allowed it; a type change keeps deciding these through the type planner.
 *
 * @module utils/chart-neutral-save-group-flags
 */

import {
	setChartBarDirection,
	setChartGrouping,
	setChartScatterStyle,
	setChartVaryColors,
	type ChartScatterStyle,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

function revert<K extends keyof PptxChartData>(
	probe: PptxChartData,
	before: PptxChartData,
	key: K,
): void {
	if (before[key] === undefined) delete probe[key];
	else probe[key] = before[key];
}

/** Plans the group settings of `probe` that differ from `before` on a kept chart type. */
export function planNeutralGroupFlags(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	// An unclaimed type change leaves the part to the object-tree path.
	if (before.chartType !== probe.chartType || probe.series.length === 0) return;
	if (before.varyColors !== probe.varyColors) {
		// The model reads the flag off the first group, which holds the first series.
		const vary = probe.varyColors ?? null;
		edits.push((space) => setChartVaryColors(space, 0, vary));
		revert(probe, before, 'varyColors');
	}
	if (
		(before.barDirection ?? 'col') !== (probe.barDirection ?? 'col') &&
		['bar', 'bar3D'].includes(probe.chartType)
	) {
		const direction = probe.barDirection ?? 'col';
		edits.push((space) => setChartBarDirection(space, direction));
		revert(probe, before, 'barDirection');
	}
	if (
		before.scatterStyle !== probe.scatterStyle &&
		probe.scatterStyle !== undefined &&
		probe.chartType === 'scatter'
	) {
		const style = probe.scatterStyle as ChartScatterStyle;
		edits.push((space) => setChartScatterStyle(space, 0, style));
		revert(probe, before, 'scatterStyle');
	}
	if (
		!!before.groupingStandard !== !!probe.groupingStandard &&
		probe.grouping === 'clustered' &&
		before.grouping === probe.grouping
	) {
		// `standard` is a 3-D bar, line and area grouping; a 2-D bar keeps `clustered`.
		const grouping = probe.groupingStandard ? 'standard' : 'clustered';
		edits.push((space) => setChartGrouping(space, grouping));
		revert(probe, before, 'groupingStandard');
	}
}
