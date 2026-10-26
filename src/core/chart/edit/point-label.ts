import type { ChartSpace } from '../model';
import type { ChartDataLabelOptions } from '../model-series';
import type { ChartManualLayout } from '../manual-layout';
import { assignOrDelete, chartSeriesEntryAt, editChartSpace } from './draft';
import { chartTitleWithText } from './title';
import { editChartTextBody, type ChartTextEdit } from './text';
import { editChartLayout } from './layout';
import { editChartShapeProperties, type ChartShapeEdit } from './shape-properties';

export interface ChartPointLabelEdit extends Omit<
	ChartDataLabelOptions,
	'spPr' | 'txPr' | 'extLst'
> {
	text?: string | null;
	font?: ChartTextEdit;
	shape?: ChartShapeEdit | null;
	layout?: ChartManualLayout | null;
}

/** Patch a point override, or remove it to inherit the series/group defaults. */
export function setChartPointLabel(
	space: ChartSpace,
	seriesIndex: number,
	pointIndex: number,
	edit: ChartPointLabelEdit | null,
): ChartSpace {
	if (!Number.isSafeInteger(pointIndex) || pointIndex < 0)
		throw new RangeError('Invalid chart label point index');
	return editChartSpace(space, (draft) => {
		const entry = chartSeriesEntryAt(draft, seriesIndex);
		if (entry.group.kind === 'surface') throw new RangeError('Surface charts have no data labels');
		const series = entry.series;
		if (!edit) {
			if (series.dataLabels)
				series.dataLabels.labels = series.dataLabels.labels.filter(
					(label) => label.index !== pointIndex,
				);
			return;
		}
		series.dataLabels ??= { labels: [] };
		let label = series.dataLabels.labels.find((item) => item.index === pointIndex);
		if (!label) {
			label = { index: pointIndex };
			series.dataLabels.labels.push(label);
		}
		const { text, font, shape, layout, ...options } = edit;
		Object.assign(label, structuredClone(options));
		if (text !== undefined) {
			if (text === null) delete label.tx;
			else {
				const title = chartTitleWithText(label.tx ? { tx: label.tx } : undefined, text, {
					preserveRuns: true,
				});
				assignOrDelete(label, 'tx', title.tx);
			}
		}
		if (font) label.txPr = editChartTextBody(label.txPr, font);
		if (shape === null) delete label.spPr;
		else if (shape) label.spPr = editChartShapeProperties(label.spPr, shape);
		if (layout !== undefined) label.layout = editChartLayout(label.layout, layout);
	});
}
