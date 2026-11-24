// The chart data sheet (`chart/data-sheet`) over a `ChartSpace`: what a product's "Edit Data"
// view shows for a chart part, and its edits mapped onto the series operations of this folder, so
// everything a series or point carries besides its data (fills, markers, labels, per-point
// formatting) is kept. Semantics follow typing in the worksheet Office opens behind the chart: a
// series name read from a cell stays a reference with a fresh cache, new series and categories
// are added the way PowerPoint's grid adds them (`Series N`, `Cat N`, zero values), and when the
// chart reads its data from the standard layout of one sheet (`workbook/layout.ts`) the formulas
// follow the grown or shrunk table, so the embedded workbook can be updated cell for cell.
import { chartSourceValues } from '../data-cache';
import {
	applyChartDataSheetEdit,
	nextChartCategoryName,
	nextChartSeriesName,
	type ChartDataSheet,
	type ChartDataSheetEdit,
} from '../data-sheet';
import type { ChartSpace } from '../model';
import {
	applyStandardChartLayout,
	standardChartFormulas,
	standardChartLayoutSheet,
} from '../workbook/layout';
import { chartSeriesEntries, chartSeriesEntryAt, editChartSpace } from './draft';
import { addChartSeries, isXYChartKind, removeChartSeries, setChartCategories } from './series';
import { addChartCategory, removeChartCategory, setChartSeriesPoint } from './series-rows';
import { categorySource } from './series-data';

/** The data sheet of `space`: the first series' categories, every series' name and values. */
export function chartSpaceDataSheet(space: ChartSpace): ChartDataSheet {
	const entries = chartSeriesEntries(space);
	const firstEntry = entries[0];
	const first = firstEntry
		? isXYChartKind(firstEntry.group.kind)
			? firstEntry.series.xValues
			: firstEntry.series.categories
		: undefined;
	const valuesOf = (index: number) => {
		const entry = entries[index]!;
		const source = isXYChartKind(entry.group.kind) ? entry.series.yValues : entry.series.values;
		return chartSourceValues(source).map((value) => (typeof value === 'number' ? value : null));
	};
	const longest = entries.reduce((max, _, index) => Math.max(max, valuesOf(index).length), 0);
	const labels = chartSourceValues(first).map(String);
	const count = Math.max(labels.length, longest);
	const categories = Array.from({ length: count }, (_, index) => labels[index] ?? '');
	return {
		categories,
		series: entries.map((entry, index) => {
			const values = valuesOf(index);
			return {
				name:
					entry.series.tx?.text ??
					entry.series.tx?.value ??
					String(chartSourceValues(entry.series.tx?.reference)[0] ?? ''),
				values: Array.from({ length: count }, (_, point) => values[point] ?? null),
			};
		}),
	};
}

/** Renames a series; a name read from a cell keeps its reference, as typing in that cell does. */
function renameSeries(space: ChartSpace, index: number, name: string): ChartSpace {
	return editChartSpace(space, (draft) => {
		const { series } = chartSeriesEntryAt(draft, index);
		const reference = series.tx?.reference;
		series.tx = reference
			? { reference: categorySource([name], reference), text: name }
			: { value: name, text: name };
	});
}

/** Category labels as the chart caches them: numbers while the cache was numeric and they parse. */
function categoryValues(space: ChartSpace, labels: readonly string[]): (string | number | null)[] {
	const entry = chartSeriesEntries(space)[0];
	const source = entry
		? isXYChartKind(entry.group.kind)
			? entry.series.xValues
			: entry.series.categories
		: undefined;
	const blank = (label: string) => label.trim() === '';
	const numeric =
		source?.cache?.type === 'number' &&
		labels.every((label) => blank(label) || Number.isFinite(Number(label)));
	return numeric ? labels.map((label) => (blank(label) ? null : Number(label))) : [...labels];
}

function applyOne(
	space: ChartSpace,
	sheet: ChartDataSheet,
	next: ChartDataSheet,
	edit: ChartDataSheetEdit,
	layoutSheet: string | undefined,
): ChartSpace {
	switch (edit.kind) {
		case 'value':
			return setChartSeriesPoint(space, edit.series, edit.category, edit.value);
		case 'series-name':
			return renameSeries(space, edit.series, edit.name);
		case 'category-label':
			return setChartCategories(space, categoryValues(space, next.categories));
		case 'add-series': {
			const index = sheet.series.length;
			const formulas = layoutSheet
				? standardChartFormulas(layoutSheet, index, next.categories.length)
				: undefined;
			return addChartSeries(space, {
				name: nextChartSeriesName(index),
				values: next.series[index]!.values,
				...(formulas ? { nameFormula: formulas.name, valuesFormula: formulas.values } : {}),
			});
		}
		case 'remove-series':
			return removeChartSeries(space, edit.series);
		case 'add-category':
			return addChartCategory(
				space,
				nextChartCategoryName(sheet.categories.length),
				0,
				sheet.categories.length,
			);
		case 'remove-category':
			return removeChartCategory(space, edit.category);
	}
}

/**
 * Applies data sheet edits to `space` in order, skipping any the sheet refuses at that point
 * (`canApplyChartDataSheetEdit`). Returns `space` itself when nothing applied.
 */
export function applyChartSpaceDataSheetEdits(
	space: ChartSpace,
	edits: readonly ChartDataSheetEdit[],
): ChartSpace {
	let sheet = chartSpaceDataSheet(space);
	const layoutSheet = standardChartLayoutSheet(space, sheet.categories.length);
	let next = space;
	for (const edit of edits) {
		const edited = applyChartDataSheetEdit(sheet, edit);
		if (!edited) continue;
		next = applyOne(next, sheet, edited, edit, layoutSheet);
		// Rows only the removed series reached go with it, as a sheet read from the part afterwards
		// shows, so the edits after it size new series and rows the same in one save as in several.
		sheet = edit.kind === 'remove-series' ? chartSpaceDataSheet(next) : edited;
	}
	if (next === space || !layoutSheet) return next;
	const count = sheet.categories.length;
	return editChartSpace(next, (draft) => applyStandardChartLayout(draft, layoutSheet, count));
}

/**
 * Chart Design > Switch Row/Column: transposes categories and series of `space`.
 * Calling it twice restores the original categories, series names and point values.
 */
export function switchChartRowColumn(space: ChartSpace): ChartSpace {
	const sheet = chartSpaceDataSheet(space);
	const oldCats = sheet.categories;
	const oldSeries = sheet.series;
	if (!oldCats.length || !oldSeries.length) return space;
	const newCategories = oldSeries.map((s) => s.name);
	const newSeriesCount = oldCats.length;
	let next = space;
	// Adjust series count to match old category count
	const currentCount = oldSeries.length;
	for (let i = currentCount; i < newSeriesCount; i++) {
		next = addChartSeries(next, {
			name: oldCats[i] ?? nextChartSeriesName(i),
			values: oldSeries.map((s) => s.values[i] ?? null),
		});
	}
	for (let i = currentCount - 1; i >= newSeriesCount; i--) {
		next = removeChartSeries(next, i);
	}
	// Adjust category count to match old series count (trimming/adding points on every series)
	for (let i = oldCats.length - 1; i >= newCategories.length; i--) {
		next = removeChartCategory(next, i);
	}
	for (let i = oldCats.length; i < newCategories.length; i++) {
		next = addChartCategory(next, newCategories[i] ?? nextChartCategoryName(i), 0, i);
	}
	// Set categories to old series names
	next = setChartCategories(next, newCategories);
	// Update each series name and point values
	for (let sIdx = 0; sIdx < newSeriesCount; sIdx++) {
		next = renameSeries(next, sIdx, oldCats[sIdx] ?? '');
		for (let cIdx = 0; cIdx < newCategories.length; cIdx++) {
			next = setChartSeriesPoint(next, sIdx, cIdx, oldSeries[cIdx]?.values[sIdx] ?? null);
		}
	}
	const layoutSheet = standardChartLayoutSheet(next, newCategories.length);
	if (layoutSheet) {
		next = editChartSpace(next, (draft) =>
			applyStandardChartLayout(draft, layoutSheet, newCategories.length),
		);
	}
	return next;
}
