import { setChartGroupDecorations, type ChartGroupDecorationsEdit } from '../../../chart';
import type { PptxChartData, PptxChartLineStyle } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';
import { neutralShapeEdit } from './chart-neutral-save-shape';

export function planNeutralChartDecorations(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	const edit: ChartGroupDecorationsEdit = {};
	const shape = (line: PptxChartLineStyle | null | undefined) => ({
		strokeColor: line?.color,
		strokeWidth: line?.width,
		strokeDashStyle: line?.dashStyle,
	});
	for (const key of ['dropLines', 'hiLowLines'] as const) {
		if (same(before[key], probe[key])) continue;
		edit[key] = probe[key] ? neutralShapeEdit(shape(before[key]), shape(probe[key])) : null;
		probe[key] = before[key];
	}
	if (!same(before.upDownBars, probe.upDownBars)) {
		const old = before.upDownBars,
			next = probe.upDownBars;
		edit.upDownBars = next
			? {
					...(old?.gapWidth !== next.gapWidth ? { gapWidth: next.gapWidth ?? null } : {}),
					...(!same(old?.upBars, next.upBars)
						? { up: neutralShapeEdit(old?.upBars, next.upBars) }
						: {}),
					...(!same(old?.downBars, next.downBars)
						? { down: neutralShapeEdit(old?.downBars, next.downBars) }
						: {}),
				}
			: null;
		probe.upDownBars = before.upDownBars;
	}
	if (Object.keys(edit).length) edits.push((space) => setChartGroupDecorations(space, edit));
}
