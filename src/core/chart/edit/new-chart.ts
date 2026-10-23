// A new chart part for a product's Insert Chart: the data sheet laid out in the standard layout of
// the embedded workbook (`workbook/layout.ts`), one clustered column group with Excel's two
// default axes (`type-axes.ts`), then switched to the picked type through `setChartType`, so a new
// chart of any family is built by the same rules as a chart changed to it. Office's defaults: the
// 4 x 3 sample Word and PowerPoint open with, a title and a legend at the bottom.
import type { ChartDataSheet } from '../data-sheet';
import type { ChartPlotGroup, ChartSpace } from '../model';
import type { ChartSpaceSeries } from '../model-series';
import { DEFAULT_CHART_SHEET_NAME, standardChartFormulas } from '../workbook/layout';
import { setChartLegend } from './legend-labels';
import { categorySource, numberSource } from './series-data';
import { setChartTitle } from './title';
import { setChartType, type ChartTypeTarget } from './type';
import { DEFAULT_CHART_AXIS_IDS, defaultChartAxes } from './type-axes';

/** The sample data Word and PowerPoint put behind a new chart. */
export const DEFAULT_CHART_DATA_SHEET: ChartDataSheet = {
	categories: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
	series: [
		{ name: 'Series 1', values: [4.3, 2.5, 3.5, 4.5] },
		{ name: 'Series 2', values: [2.4, 4.4, 1.8, 2.8] },
		{ name: 'Series 3', values: [2, 2, 3, 5] },
	],
};

export interface NewChartOptions {
	/** The chart title; absent for no title (`c:autoTitleDeleted`). */
	title?: string;
	/** The worksheet the formulas name; Office's `Sheet1` by default. */
	sheetName?: string;
}

function series(sheet: ChartDataSheet, index: number, sheetName: string): ChartSpaceSeries {
	const entry = sheet.series[index]!;
	const formulas = standardChartFormulas(sheetName, index, sheet.categories.length);
	return {
		index,
		order: index,
		tx: { reference: categorySource([entry.name], undefined, formulas.name), text: entry.name },
		invertIfNegative: false,
		dataPoints: [],
		categories: categorySource(sheet.categories, undefined, formulas.categories),
		values: numberSource(entry.values, undefined, formulas.values),
	};
}

/** A new chart part showing `sheet` as `target` (a column chart when `target` is absent). */
export function newChartSpace(
	sheet: ChartDataSheet = DEFAULT_CHART_DATA_SHEET,
	target?: ChartTypeTarget,
	options: NewChartOptions = {},
): ChartSpace {
	const sheetName = options.sheetName ?? DEFAULT_CHART_SHEET_NAME;
	const group: ChartPlotGroup = {
		kind: 'bar',
		element: 'barChart',
		is3D: false,
		barDirection: 'col',
		grouping: 'clustered',
		varyColors: false,
		gapWidth: 219,
		overlap: -27,
		series: sheet.series.map((_, index) => series(sheet, index, sheetName)),
		axisIds: [...DEFAULT_CHART_AXIS_IDS],
	};
	let space: ChartSpace = {
		roundedCorners: false,
		autoTitleDeleted: options.title === undefined,
		plotArea: { layout: {}, groups: [group], axes: defaultChartAxes(false, false) },
		plotVisibleOnly: true,
		displayBlanksAs: 'gap',
	};
	const column = !target || (target.kind === 'bar' && (target.barDirection ?? 'col') === 'col');
	if (target && !(column && !target.is3D && (target.grouping ?? 'clustered') === 'clustered'))
		space = setChartType(space, target);
	if (options.title !== undefined) space = setChartTitle(space, options.title);
	return setChartLegend(space, { position: 'b' });
}
