import type { ChartDataSheet, ChartDataSheetEdit } from 'ooxml-core/chart';

/**
 * Translated text of the sheet. `{number}`, `{name}`, `{series}` and `{categories}` in a template
 * (or the i18next form, `{{number}}`) are replaced with the cell's position, series name and the
 * counts.
 */
export interface OfficeChartDataSheetLabels {
	heading: string;
	addCategory: string;
	addSeries: string;
	addCategoryShort: string;
	addSeriesShort: string;
	removeCategory: string;
	removeSeries: string;
	categories: string;
	seriesName: string;
	categoryLabel: string;
	valueCell: string;
	summary: string;
	invalidNumber: string;
}

export const CHART_DATA_SHEET_ENGLISH: OfficeChartDataSheetLabels = {
	heading: 'Data',
	addCategory: 'Add category',
	addSeries: 'Add series',
	addCategoryShort: 'Cat',
	addSeriesShort: 'Series',
	removeCategory: 'Remove category',
	removeSeries: 'Remove series',
	categories: 'Categories',
	seriesName: 'Series {number}',
	categoryLabel: 'Categories {number}',
	valueCell: '{name} value {number}',
	summary: '{series} series · {categories} categories',
	invalidNumber: 'Enter a number',
};

/** The cell the host marks (a chart part picked on the canvas): a value, or a whole series. */
export interface OfficeChartDataSheetHighlight {
	series: number;
	point?: number;
}

/** Detail of `office-chart-data-change`: the edits in order, and the sheet after them. */
export interface OfficeChartDataChange {
	edits: ChartDataSheetEdit[];
	data: ChartDataSheet;
}
export type OfficeChartDataChangeEvent = CustomEvent<OfficeChartDataChange>;

/** Fills a label template's `{key}` (or `{{key}}`) placeholders. */
export function fill(template: string, values: Record<string, string | number>): string {
	return template.replace(/\{\{?(\w+)\}?\}/g, (whole, key: string) =>
		key in values ? String(values[key]) : whole,
	);
}
