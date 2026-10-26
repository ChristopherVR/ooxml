// Word's ChartEx adapter uses the same neutral operations and image-node history as classic charts.
import {
	applyChartExSpaceDataSheetEdits,
	chartExSpaceDataSheet,
	chartExDataChanged,
	chartExSpaceWorkbookWrites,
	chartExSummaryFromModel,
	type ChartDataSheetEdit,
	type ChartExSpace,
} from '../chart';
import type { DocxChart } from './chart';

export function isDocxChartExEditable(
	chart: DocxChart | undefined,
): chart is DocxChart & { chartExSpace: ChartExSpace } {
	return (
		!!chart?.chartExSpace &&
		chart.issues.every((issue) => issue.code === 'CHARTEX_FAMILY_NOT_DRAWN')
	);
}
export function editDocxChartEx(
	chart: DocxChart,
	operation: (space: ChartExSpace) => ChartExSpace,
): DocxChart {
	if (!isDocxChartExEditable(chart)) throw new Error('This extended chart cannot be edited.');
	const space = operation(chart.chartExSpace);
	if (space === chart.chartExSpace) return chart;
	if (chartExDataChanged(chart.chartExSpace, space)) {
		if (space.chartData.externalData && !chart.workbookPartName)
			throw new Error('Extended chart data cannot update an external or missing workbook.');
		const writes = chartExSpaceWorkbookWrites(space);
		if (writes.length && !chart.workbookPartName)
			throw new Error('Extended chart data has no embedded workbook.');
		if (
			space.chartData.externalData &&
			space.plotArea.series.some(
				(series) =>
					!series.tx?.formula ||
					(
						series.data ?? space.chartData.data.find((data) => data.id === series.dataId)
					)?.dimensions.some((d) => !d.formula),
			)
		)
			throw new Error('New series require a mapped workbook range.');
	}
	const { chartEx: _summary, title: _title, ...rest } = chart;
	const summary = chartExSummaryFromModel(space);
	const title = space.title?.tx?.value ?? space.title?.tx?.rich?.text;
	return {
		...rest,
		chartExSpace: space,
		layoutId: summary.layoutId,
		...(summary.summary ? { chartEx: summary.summary } : {}),
		...(title ? { title } : {}),
	};
}
export function docxChartExDataSheet(chart: DocxChart | undefined) {
	return isDocxChartExEditable(chart) ? chartExSpaceDataSheet(chart.chartExSpace) : undefined;
}
export function applyDocxChartExDataSheetEdits(
	chart: DocxChart,
	edits: readonly ChartDataSheetEdit[],
): DocxChart {
	return editDocxChartEx(chart, (space) => applyChartExSpaceDataSheetEdits(space, edits));
}
