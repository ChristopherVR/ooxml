/**
 * The chart data sheet (`chart/data-sheet`) over a `ChartSpace`: reading a chart part's cached data
 * as the sheet, and applying the sheet's edits through the series operations of `chart/edit`, so
 * formatting on series and points stays with the part. Scatter and bubble series read and write x
 * values as categories and y values as values (`chart/edit` series semantics).
 *
 * @module chart/data-sheet-space
 */
import { chartSourceValues } from './data-cache';
import {
	applyChartDataSheetEdit,
	nextChartCategoryName,
	nextChartSeriesName,
	type ChartDataSheet,
	type ChartDataSheetEdit,
} from './data-sheet';
import { chartSeriesEntries } from './edit/draft';
import {
	addChartCategory,
	addChartSeries,
	isXYChartKind,
	removeChartCategory,
	removeChartSeries,
	renameChartSeries,
	setChartCategories,
	setChartSeriesValues,
} from './edit/series';
import type { ChartSpace } from './model';

const seriesName = (text: string | undefined, index: number) => text ?? `Series ${index + 1}`;

/** The data sheet of `space`: the first series' categories and every series' name and values. */
export function chartSpaceDataSheet(space: ChartSpace): ChartDataSheet {
	const entries = chartSeriesEntries(space);
	const first = entries[0];
	const categories = first
		? chartSourceValues(
				isXYChartKind(first.group.kind) ? first.series.xValues : first.series.categories,
			).map(String)
		: [];
	const series = entries.map(({ group, series: source }, index) => {
		const raw = chartSourceValues(isXYChartKind(group.kind) ? source.yValues : source.values);
		const values = categories.map((_, i) => {
			const value = raw[i];
			return typeof value === 'number' ? value : null;
		});
		const name =
			source.tx?.text ?? source.tx?.value ?? chartSourceValues(source.tx?.reference)[0]?.toString();
		return { name: seriesName(name, index), values };
	});
	return { categories, series };
}

/**
 * Applies one data sheet edit to `space` through `chart/edit`, or returns `space` unchanged when the
 * sheet refuses it (`canApplyChartDataSheetEdit`). New series and categories take PowerPoint's
 * names (`Series N`, `Cat N`) and zero values.
 */
export function applyChartDataSheetEditToSpace(
	space: ChartSpace,
	edit: ChartDataSheetEdit,
): ChartSpace {
	const sheet = chartSpaceDataSheet(space);
	const next = applyChartDataSheetEdit(sheet, edit);
	if (!next) return space;
	switch (edit.kind) {
		case 'value':
			return setChartSeriesValues(space, edit.series, next.series[edit.series]!.values);
		case 'series-name':
			return renameChartSeries(space, edit.series, edit.name);
		case 'category-label':
			return setChartCategories(space, next.categories);
		case 'add-series':
			return addChartSeries(space, {
				name: nextChartSeriesName(sheet.series.length),
				values: sheet.categories.map(() => 0),
			});
		case 'remove-series':
			return removeChartSeries(space, edit.series);
		case 'add-category':
			return addChartCategory(space, nextChartCategoryName(sheet.categories.length), 0);
		case 'remove-category':
			return removeChartCategory(space, edit.category);
	}
}

/** Applies edits in order (see {@link applyChartDataSheetEditToSpace}). */
export function applyChartDataSheetEditsToSpace(
	space: ChartSpace,
	edits: readonly ChartDataSheetEdit[],
): ChartSpace {
	return edits.reduce(applyChartDataSheetEditToSpace, space);
}
