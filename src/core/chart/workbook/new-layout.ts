/**
 * The data of a new chart moved into the embedded workbook it is created with: every series name,
 * category and value source becomes a cell reference (`c:strRef`, `c:numRef`) that keeps its
 * cache, and the worksheet rows hold the same values, so "Edit Data" opens the chart's numbers.
 *
 * A chart whose series share one category vector is laid out in the standard layout
 * (`layout.ts`): categories in column A, names across row 1, values under them, which the data
 * sheet edits grow and shrink. Scatter and bubble charts, and series with categories of their own,
 * take one column per vector (`rebase-data.ts`: X values, Y values, bubble sizes). A blank point
 * (no cached point) is an empty cell.
 *
 * @module chart/workbook/new-layout
 */
import { chartSourceValues } from '../data-cache';
import { chartSeriesEntries } from '../edit/draft';
import { isXYChartKind } from '../edit/series';
import type { ChartSpace } from '../model';
import type { ChartDataCache, ChartDataSource } from '../model-series';
import { DEFAULT_CHART_SHEET_NAME, standardChartFormulas } from './layout';
import { rebaseChartData } from './rebase-data';

/** A new chart's data in its workbook: the referencing chart and the worksheet rows. */
export interface ChartWorkbookLayout {
	chartSpace: ChartSpace;
	/** Worksheet rows from row 1; `null` is an empty cell. */
	rows: (string | number | null)[][];
}

/** The cells of a cache, `count` long: numbers in a number cache, `null` for a missing point. */
function cacheCells(cache: ChartDataCache | undefined, count: number): (string | number | null)[] {
	const cells: (string | number | null)[] = Array.from({ length: count }, () => null);
	for (const [at, point] of (cache?.points ?? []).entries()) {
		const index = point.index ?? at;
		if (index < 0 || index >= count || point.value === '') continue;
		const number = Number(point.value);
		cells[index] =
			cache?.type === 'number' && point.value.trim() !== '' && Number.isFinite(number)
				? number
				: point.value;
	}
	return cells;
}

const pointCount = (source: ChartDataSource | undefined) =>
	Math.max(source?.cache?.pointCount ?? 0, chartSourceValues(source).length);

/** Whether every series reads the same flat categories (or none), so column A can hold them. */
function sharesCategories(space: ChartSpace): boolean {
	const entries = chartSeriesEntries(space);
	if (entries.some(({ group }) => isXYChartKind(group.kind))) return false;
	if (entries.some(({ series }) => series.categories?.levels || !series.values)) return false;
	const first = JSON.stringify(chartSourceValues(entries[0]?.series.categories));
	return entries.every(
		({ series }) => JSON.stringify(chartSourceValues(series.categories)) === first,
	);
}

/** The standard layout: categories in column A, each series' name in row 1 over its values. */
function standardLayout(space: ChartSpace, sheetName: string): ChartWorkbookLayout {
	const chartSpace = structuredClone(space);
	const entries = chartSeriesEntries(chartSpace);
	const categories = entries[0]?.series.categories;
	const count = Math.max(
		1,
		pointCount(categories),
		...entries.map(({ series }) => pointCount(series.values)),
	);
	const rows: (string | number | null)[][] = [[null]];
	cacheCells(categories?.cache, count).forEach((cell, index) => (rows[index + 1] = [cell]));
	entries.forEach(({ series }, index) => {
		const formulas = standardChartFormulas(sheetName, index, count);
		const name = series.tx?.text ?? series.tx?.value ?? `Series ${index + 1}`;
		rows[0]![index + 1] = name;
		series.tx = {
			text: name,
			reference: {
				kind: 'strRef',
				formula: formulas.name,
				cache: { type: 'string', pointCount: 1, points: [{ index: 0, value: name }] },
			},
		};
		if (series.categories) {
			series.categories.formula = formulas.categories;
			series.categories.kind = series.categories.cache?.type === 'number' ? 'numRef' : 'strRef';
		}
		const values = series.values!;
		values.formula = formulas.values;
		values.kind = 'numRef';
		cacheCells(values.cache, count).forEach((cell, point) => (rows[point + 1]![index + 1] = cell));
	});
	return { chartSpace, rows };
}

/**
 * `space` (literal data) pointed at `sheetName` of a new embedded workbook, with that sheet's
 * rows. Throws a `RangeError` when a source cannot be placed (a series without values).
 */
export function newChartWorkbookLayout(
	space: ChartSpace,
	sheetName: string = DEFAULT_CHART_SHEET_NAME,
): ChartWorkbookLayout {
	return sharesCategories(space)
		? standardLayout(space, sheetName)
		: rebaseChartData(space, sheetName);
}
