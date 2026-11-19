/**
 * The embedded workbook as the fallback of a chart part's empty caches: a chart saved without
 * caches (COM automation that closes the workbook right after `SetSourceData`, some generators)
 * still carries its data in the embedded workbook, which the part parser never opens.
 *
 * @module runtime/chart-embedded-fallback
 */
import type { ChartDataSource, ChartPlotGroup } from '../../../../chart';
import type { PptxChartData } from '../../types';

const uncached = (source: ChartDataSource | undefined) =>
	!!source?.formula && !source.levels && !source.cache?.points.length;

/**
 * Whether the part names its categories (or, scatter and bubble, its x values) in the workbook
 * without caching them. A part that names none has no category labels to read: its first column
 * may be anything, a series of its own included.
 */
export const hasUncachedCategoryReference = (groups: readonly ChartPlotGroup[]): boolean =>
	groups.some(({ series: [first] }) => uncached(first?.categories) || uncached(first?.xValues));

/**
 * Whether {@link applyEmbeddedWorkbookFallback} can change anything: the part names categories it
 * does not cache, caches no series values, or (scatter and bubble) no x values for some series.
 */
export function needsEmbeddedWorkbook(
	chartType: PptxChartData['chartType'],
	categories: readonly string[],
	series: PptxChartData['series'],
	categoriesUncached: boolean,
): boolean {
	return (
		(categories.length === 0 && categoriesUncached) ||
		series.every((entry) => entry.values.length === 0) ||
		((chartType === 'scatter' || chartType === 'bubble') &&
			series.some((entry) => !entry.xValues || entry.xValues.length === 0))
	);
}

/**
 * Fills what the part does not cache from the embedded workbook's rows: the categories when the
 * part names categories it does not cache, every series' values when all are empty, and the x
 * values of scatter and bubble series that cache none (the worksheet reader puts that column in
 * `categories`).
 */
export function applyEmbeddedWorkbookFallback(
	chartType: PptxChartData['chartType'],
	categories: string[],
	series: PptxChartData['series'],
	workbook: PptxChartData['embeddedWorkbookData'],
	categoriesUncached: boolean,
): { categories: string[]; series: PptxChartData['series'] } {
	if (!workbook) {
		return { categories, series };
	}
	let finalCategories = categories;
	let finalSeries = series;
	if (finalCategories.length === 0 && categoriesUncached && workbook.categories.length > 0) {
		finalCategories = workbook.categories;
	}
	const allSeriesEmpty = finalSeries.every((entry) => entry.values.length === 0);
	if (allSeriesEmpty && workbook.series.length > 0) {
		finalSeries = finalSeries.map((entry, index) => {
			const values = workbook.series[index]?.values;
			return values && values.length > 0 ? { ...entry, values } : entry;
		});
	}
	if ((chartType === 'scatter' || chartType === 'bubble') && workbook.categories.length > 0) {
		const xValues = workbook.categories.map(Number);
		if (xValues.every(Number.isFinite)) {
			finalSeries = finalSeries.map((entry) =>
				entry.xValues && entry.xValues.length > 0 ? entry : { ...entry, xValues },
			);
		}
	}
	return { categories: finalCategories, series: finalSeries };
}
