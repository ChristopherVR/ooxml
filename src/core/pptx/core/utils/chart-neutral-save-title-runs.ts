// Explicit run edits use neutral run properties; flat title edits are handled by the existing planner.
import { setChartTitleRuns, type ChartTitleRunEdit } from '../../../chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';
import { neutralTextEdit } from './chart-neutral-save-text';
export function planNeutralTitleRuns(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	if (
		same(before.titleRuns, probe.titleRuns) ||
		!probe.titleRuns?.length ||
		probe.style?.hasTitle === false
	)
		return;
	// The host's title field names the first run; a diverged flat title still uses positional realignment.
	if (probe.title !== undefined && probe.title !== probe.titleRuns[0]!.text) return;
	const runs: ChartTitleRunEdit[] = probe.titleRuns.map((run, index) => {
		const old = before.titleRuns?.[index] ?? { text: '' };
		const format = neutralTextEdit(
			{ fontSize: old.fontSize, bold: old.bold, italic: old.italic, color: old.color },
			{ fontSize: run.fontSize, bold: run.bold, italic: run.italic, color: run.color },
		);
		const inherit = (['bold', 'italic'] as const).filter(
			(key) => old[key] !== undefined && run[key] === undefined,
		);
		return { text: run.text, format, ...(inherit.length ? { inherit } : {}) };
	});
	edits.push((space) => setChartTitleRuns(space, runs));
	probe.titleRuns = before.titleRuns;
	probe.title = before.title;
}
