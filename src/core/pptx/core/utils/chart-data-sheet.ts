/**
 * Adapters between a pptx chart's data (`PptxChartData`) and the format-neutral chart data sheet
 * (`chart/data-sheet`), the table every binding's "Edit Data" view shows.
 *
 * Reading is a projection: category labels, series names and values. Writing goes edit by edit
 * through the chart grid operations, never by rebuilding series from the sheet, so everything a
 * series or a point carries besides its values (colours, markers, per-point formatting, labels,
 * trendlines) is kept, and removing a category re-indexes per-point formatting exactly as the
 * grid always has (`followDataPoint`).
 *
 * @module utils/chart-data-sheet
 */
import type { ChartDataSheet, ChartDataSheetEdit } from '../../../chart/data-sheet';
import type { PptxChartData } from '../types/chart';
import {
	chartGridAddCategory,
	chartGridAddSeries,
	chartGridRemoveCategory,
	chartGridRemoveSeries,
	chartGridSetCategoryLabel,
	chartGridSetCellValue,
} from './chart-grid-operations';

/** The sheet a chart's data shows: a series shorter than the categories reads as blank cells. */
export function chartDataToSheet(data: PptxChartData): ChartDataSheet {
	const count = data.categories.length;
	return {
		categories: [...data.categories],
		series: data.series.map((series) => ({
			name: series.name,
			values: Array.from({ length: count }, (_, index) => {
				const value = series.values[index];
				return typeof value === 'number' && Number.isFinite(value) ? value : null;
			}),
		})),
	};
}

/** Pads a ragged series (legal in OOXML) with zeroes up to `index` so its cell can be written. */
function padSeries(data: PptxChartData, seriesIndex: number, index: number): PptxChartData {
	const series = data.series[seriesIndex];
	if (!series || index < series.values.length) {
		return data;
	}
	const values = [...series.values];
	while (values.length <= index) {
		values.push(0);
	}
	return {
		...data,
		series: data.series.map((entry, i) => (i === seriesIndex ? { ...entry, values } : entry)),
	};
}

/** One sheet edit on a chart's data, or `null` when the grid policy refuses it. */
export function chartDataApplySheetEdit(
	data: PptxChartData,
	edit: ChartDataSheetEdit,
	followDataPoint = true,
): PptxChartData | null {
	switch (edit.kind) {
		case 'value':
			if (edit.category < 0 || edit.category >= data.categories.length) {
				return null;
			}
			return chartGridSetCellValue(
				padSeries(data, edit.series, edit.category),
				edit.series,
				edit.category,
				String(edit.value),
			);
		case 'series-name':
			if (!Number.isInteger(edit.series) || !data.series[edit.series]) {
				return null;
			}
			return {
				...data,
				series: data.series.map((entry, index) =>
					index === edit.series ? { ...entry, name: edit.name } : entry,
				),
			};
		case 'category-label':
			return chartGridSetCategoryLabel(data, edit.category, edit.label);
		case 'add-series':
			return chartGridAddSeries(data);
		case 'remove-series':
			return chartGridRemoveSeries(data, edit.series);
		case 'add-category':
			return chartGridAddCategory(data);
		case 'remove-category':
			return chartGridRemoveCategory(data, edit.category, followDataPoint);
		default:
			return null;
	}
}

/**
 * Sheet edits applied in order (a refused edit is skipped), or `null` when none applied, so a
 * caller commits one history entry per sheet change: `const next = ...; if (next) commit(next)`.
 */
export function chartDataApplySheetEdits(
	data: PptxChartData,
	edits: readonly ChartDataSheetEdit[],
	followDataPoint = true,
): PptxChartData | null {
	let next = data;
	let applied = false;
	for (const edit of edits) {
		const result = chartDataApplySheetEdit(next, edit, followDataPoint);
		if (result) {
			next = result;
			applied = true;
		}
	}
	return applied ? next : null;
}
