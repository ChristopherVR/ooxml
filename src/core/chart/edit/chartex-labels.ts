import type { ChartExSpace, ChartExLabelVisibility, ChartNumberFormat } from '../index';
import { chartExSeriesAt, editChartExSpace } from './chartex';
import { editChartTextBody, type ChartTextEdit } from './text';
import { editChartShapeProperties, type ChartShapeEdit } from './shape-properties';
export interface ChartExLabelsEdit {
	visibility?: Partial<Record<keyof ChartExLabelVisibility, boolean | null>>;
	position?: string | null;
	separator?: string | null;
	numberFormat?: ChartNumberFormat | null;
	font?: ChartTextEdit;
	shape?: ChartShapeEdit | null;
}
/** The group defaults are patched without replacing point labels, hidden points or source XML. */
export function editChartExLabels(
	space: ChartExSpace,
	seriesIndex: number | undefined,
	edit: ChartExLabelsEdit,
): ChartExSpace {
	return editChartExSpace(space, (draft) => {
		const series =
			seriesIndex === undefined ? draft.plotArea.series : [chartExSeriesAt(draft, seriesIndex)];
		for (const item of series) {
			const group = (item.dataLabels ??= { labels: [], hidden: [] });
			if (edit.visibility) {
				group.visibility ??= {};
				for (const key of Object.keys(edit.visibility) as (keyof ChartExLabelVisibility)[]) {
					const value = edit.visibility[key];
					if (value === null) delete group.visibility[key];
					else if (value !== undefined) group.visibility[key] = value;
				}
			}
			for (const key of ['position', 'separator', 'numberFormat'] as const)
				if (edit[key] !== undefined) {
					if (edit[key] === null) delete group[key];
					else Object.assign(group, { [key]: structuredClone(edit[key]) });
				}
			if (edit.font) group.txPr = editChartTextBody(group.txPr, edit.font);
			if (edit.shape === null) delete group.spPr;
			else if (edit.shape) group.spPr = editChartShapeProperties(group.spPr, edit.shape);
		}
	});
}
