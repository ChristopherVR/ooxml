import { setChartLayout } from '../../../chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';

export function planNeutralChartLayout(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	for (const region of ['title', 'legend', 'plotArea'] as const) {
		const old = before.layouts?.[region],
			value = probe.layouts?.[region];
		if (same(old, value)) continue;
		// A changed raw extension needs the source-tree adapter, not a lossy conversion.
		if (!same(old?.ext, value?.ext)) continue;
		const { ext: _ext, ...layout } = value ?? {};
		edits.push((space) => setChartLayout(space, region, value == null ? null : layout));
		probe.layouts ??= {};
		if (old === undefined) delete probe.layouts[region];
		else probe.layouts[region] = old;
	}
	if (probe.layouts && !Object.keys(probe.layouts).length && before.layouts === undefined)
		delete probe.layouts;
}
