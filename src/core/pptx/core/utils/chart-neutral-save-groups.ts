import {
	chartSeriesEntries,
	setChartGroupOptions,
	type ChartGroupOptionsEdit,
} from '../../../chart/index';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';

const OPTIONS = {
	barGapWidth: 'gapWidth',
	barOverlap: 'overlap',
	firstSliceAngle: 'firstSliceAngle',
	doughnutHoleSize: 'holeSize',
} as const;

/** Compatibility field mapping; the shared operation validates group schema and ranges. */
export function planNeutralGroupOptions(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	for (const key of Object.keys(OPTIONS) as (keyof typeof OPTIONS)[]) {
		if (before[key] === probe[key]) continue;
		const edit: ChartGroupOptionsEdit = { [OPTIONS[key]]: probe[key] ?? null };
		edits.push((space) => {
			const visited = new Set<number>();
			let next = space;
			for (const [index, { group, groupIndex }] of chartSeriesEntries(space).entries()) {
				if (visited.has(groupIndex)) continue;
				visited.add(groupIndex);
				const allowed =
					key === 'barGapWidth'
						? ['barChart', 'bar3DChart']
						: key === 'barOverlap'
							? ['barChart']
							: key === 'firstSliceAngle'
								? ['pieChart', 'doughnutChart']
								: ['doughnutChart'];
				if (allowed.includes(group.element)) next = setChartGroupOptions(next, index, edit);
			}
			return next;
		});
		if (before[key] === undefined) delete probe[key];
		else probe[key] = before[key];
	}
}
