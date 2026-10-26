// Patch supported point-label metadata. ChartEx has no custom text or manual-label layout node.
import type { ChartExSpace, ChartExDataLabelOptions } from '../ex/model';
import { chartExSeriesAt, editChartExSpace } from './chartex';
import { editChartTextBody, type ChartTextEdit } from './text';
import { editChartShapeProperties, type ChartShapeEdit } from './shape-properties';
export interface ChartExPointLabelEdit extends Omit<
	ChartExDataLabelOptions,
	'spPr' | 'txPr' | 'extLst'
> {
	deleted?: boolean;
	font?: ChartTextEdit;
	shape?: ChartShapeEdit | null;
}
export function editChartExPointLabel(
	space: ChartExSpace,
	seriesIndex: number,
	pointIndex: number,
	edit: ChartExPointLabelEdit | null,
): ChartExSpace {
	if (!Number.isSafeInteger(pointIndex) || pointIndex < 0)
		throw new RangeError('Invalid point index.');
	return editChartExSpace(space, (draft) => {
		const series = chartExSeriesAt(draft, seriesIndex);
		series.dataLabels ??= { labels: [], hidden: [] };
		const group = series.dataLabels;
		if (!edit) {
			group.labels = group.labels.filter((label) => label.index !== pointIndex);
			group.hidden = group.hidden.filter((index) => index !== pointIndex);
			return;
		}
		let label = group.labels.find((label) => label.index === pointIndex);
		if (!label) {
			label = { index: pointIndex };
			group.labels.push(label);
		}
		const { deleted, font, shape, visibility, ...rest } = edit;
		Object.assign(label, structuredClone(rest));
		if (visibility) label.visibility = { ...label.visibility, ...visibility };
		if (deleted !== undefined) {
			group.hidden = group.hidden.filter((index) => index !== pointIndex);
			if (deleted) group.hidden.push(pointIndex);
		}
		if (font) label.txPr = editChartTextBody(label.txPr, font);
		if (shape === null) delete label.spPr;
		else if (shape) label.spPr = editChartShapeProperties(label.spPr, shape);
	});
}
