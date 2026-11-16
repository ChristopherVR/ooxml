// Word's ChartEx adapter uses the same neutral operations and image-node history as classic charts.
import {
	applyChartExSpaceDataSheetEdits,
	chartExSpaceDataSheet,
	chartExDataChanged,
	chartExSpaceWorkbookWrites,
	chartExSummaryFromModel,
	applyChartExOfficeDefaults,
	createChartExSpace,
	type ChartDataSheet,
	type ChartDataSheetEdit,
	type ChartExDataSheet,
	type ChartExFamily,
	type ChartExSpace,
	type CreateChartExOptions,
} from '../chart';
import { CHART_NOTICE, type DocxChart } from './chart';

/** Office's sample data for a new extended chart of `family` (Insert > Chart in Word). */
export function chartExSampleSheet(family: ChartExFamily): ChartDataSheet | ChartExDataSheet {
	switch (family) {
		case 'regionMap':
			return {
				categories: ['United States', 'Canada', 'Brazil', 'France', 'India'],
				series: [{ name: 'Series 1', values: [9, 7, 5, 3, 6] }],
			};
		case 'pareto':
			return {
				categories: ['Category 1', 'Category 2', 'Category 3', 'Category 4', 'Category 5'],
				series: [{ name: 'Series 1', values: [45, 30, 12, 8, 5] }],
			};
		case 'waterfall':
			return {
				categories: ['Start', 'Q1', 'Q2', 'Q3', 'End'],
				series: [{ name: 'Series 1', values: [100, 20, -15, 30, 135] }],
			};
		case 'funnel':
			return {
				categories: ['Stage 1', 'Stage 2', 'Stage 3', 'Stage 4'],
				series: [{ name: 'Series 1', values: [100, 70, 45, 20] }],
			};
		case 'treemap':
		case 'sunburst': {
			const leaves = ['Leaf 1', 'Leaf 2', 'Leaf 3', 'Leaf 4', 'Leaf 5', 'Leaf 6'];
			return {
				categories: leaves,
				categoryLevels: [leaves, ['Branch 1', '', '', 'Branch 2', '', '']],
				series: [{ name: 'Series 1', values: [22, 12, 18, 87, 88, 17] }],
			};
		}
		case 'boxWhisker':
			return {
				categories: [
					'Group 1',
					'Group 1',
					'Group 1',
					'Group 1',
					'Group 2',
					'Group 2',
					'Group 2',
					'Group 2',
				],
				series: [{ name: 'Series 1', values: [12, 18, 25, 30, 8, 14, 21, 40] }],
			};
		default:
			return {
				categories: ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'],
				series: [{ name: 'Series 1', values: [3, 5, 8, 4, 6, 7, 12, 9] }],
			};
	}
}

/**
 * A new extended chart for Insert > Chart: `sheet` (or the sample data) as `family`. It has no part
 * yet; saving writes the `chartEx` part, its relationships and an embedded workbook.
 */
export function newDocxChartEx(
	family: ChartExFamily,
	options: CreateChartExOptions & { sheet?: ChartDataSheet | ChartExDataSheet } = {},
): DocxChart {
	const { sheet, ...create } = options;
	const space = applyChartExOfficeDefaults(
		createChartExSpace(sheet ?? chartExSampleSheet(family), family, create),
		family,
	);
	const summary = chartExSummaryFromModel(space);
	const title = space.title?.tx?.value ?? space.title?.tx?.rich?.text;
	return {
		placement: 'inline',
		notice: CHART_NOTICE,
		issues: [],
		chartExSpace: space,
		layoutId: summary.layoutId,
		...(summary.summary ? { chartEx: summary.summary } : {}),
		...(title ? { title } : {}),
	};
}

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
	let rebase = chart.chartExRebaseData ?? false;
	if (space === chart.chartExSpace) return chart;
	if (chartExDataChanged(chart.chartExSpace, space)) {
		if (space.chartData.externalData && !chart.workbookPartName)
			throw new Error('Extended chart data cannot update an external or missing workbook.');
		if (!rebase) {
			try {
				const writes = chartExSpaceWorkbookWrites(space);
				if (writes.length && !chart.workbookPartName)
					throw new Error('Extended chart data has no embedded workbook.');
				if (
					space.plotArea.series.some(
						(series) =>
							// A pareto line has neither name nor data of its own.
							(series.ownerIndex === undefined && !series.tx?.formula) ||
							(
								series.data ?? space.chartData.data.find((data) => data.id === series.dataId)
							)?.dimensions.some((d) => !d.formula),
					)
				)
					rebase = true;
			} catch (error) {
				// Only a local embedded source may grow into a dedicated new worksheet.
				chartExSpaceWorkbookWrites(chart.chartExSpace);
				if (!chart.workbookPartName) throw error;
				rebase = true;
			}
		}
	}
	const { chartEx: _summary, title: _title, ...rest } = chart;
	const summary = chartExSummaryFromModel(space);
	const title = space.title?.tx?.value ?? space.title?.tx?.rich?.text;
	return {
		...rest,
		chartExSpace: space,
		...(rebase ? { chartExRebaseData: true } : {}),
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
