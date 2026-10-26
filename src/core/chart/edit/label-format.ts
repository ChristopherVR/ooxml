import type { ChartSpace } from '../model';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { editChartShapeProperties, type ChartShapeEdit } from './shape-properties';
import { editChartTextBody, type ChartTextEdit } from './text';

export interface ChartDataLabelFormattingEdit {
	font?: ChartTextEdit;
	shape?: ChartShapeEdit;
	showLeaderLines?: boolean | null;
	leaderLine?: ChartShapeEdit | null;
}

/** Patch defaults at one label level without overwriting individual point overrides. */
export function setChartDataLabelFormatting(
	space: ChartSpace,
	edit: ChartDataLabelFormattingEdit,
	target: { series?: number } = {},
): ChartSpace {
	return editChartSpace(space, (draft) => {
		const holders =
			target.series === undefined
				? draft.plotArea.groups.filter((group) => group.kind !== 'surface')
				: [chartSeriesEntryAt(draft, target.series).series];
		for (const holder of holders) {
			const labels = (holder.dataLabels ??= { labels: [] });
			if (edit.font) labels.txPr = editChartTextBody(labels.txPr, edit.font);
			if (edit.shape) labels.spPr = editChartShapeProperties(labels.spPr, edit.shape);
			if (edit.showLeaderLines === null) delete labels.showLeaderLines;
			else if (edit.showLeaderLines !== undefined) labels.showLeaderLines = edit.showLeaderLines;
			if (edit.leaderLine === null) delete labels.leaderLines;
			else if (edit.leaderLine) {
				labels.leaderLines ??= {};
				labels.leaderLines.spPr = editChartShapeProperties(
					labels.leaderLines.spPr,
					edit.leaderLine,
				);
			}
		}
	});
}
