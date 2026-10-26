// Reuse the classic neutral DrawingML planners; only the target style holders cross the bridge.
import { editChartExStyle } from '../../../chart';
import type { PptxChartData } from '../types';
import type { ChartExEdit } from './chart-ex-neutral-save';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { planNeutralChartFormats } from './chart-neutral-save-format';
import { planNeutralChartText } from './chart-neutral-save-text';
import { planNeutralChartShapes } from './chart-neutral-save-shape';
import { planNeutralFlatPaint } from './chart-neutral-save-flat-paint';
import { same } from './chart-neutral-save-common';
export function planChartExStyles(
	before: PptxChartData,
	probe: PptxChartData,
	edits: ChartExEdit[],
): void {
	const shared: NeutralChartEdit[] = [];
	planNeutralChartFormats(before, probe, shared);
	// ChartEx has no per-entry legend override node: leave that field honestly unclaimed.
	const entries = probe.style?.legendEntries;
	const changed = !same(entries, before.style?.legendEntries);
	if (changed && probe.style) probe.style.legendEntries = before.style?.legendEntries;
	planNeutralChartText(before, probe, shared);
	if (changed && probe.style) probe.style.legendEntries = entries;
	planNeutralChartShapes(before, probe, shared);
	planNeutralFlatPaint(before, probe, shared);
	for (const edit of shared) edits.push((space) => editChartExStyle(space, edit));
}
