/**
 * Pure chart data manipulation utilities.
 *
 * These functions operate on {@link PptxChartData} directly (immutable style,
 * returning new objects) and are framework-agnostic. They complement the
 * element-level SDK operations in `chart-operations.ts` which mutate
 * {@link ChartPptxElement} in place.
 *
 * Naming convention: `chartData*` prefix to avoid collisions with the
 * element-level SDK operations (`addChartSeries`, etc.).
 *
 * @module utils/chart-data-utils
 */

import type {
	PptxChartData,
	PptxChartDataPoint,
	PptxChartSeries,
	PptxChartStructureEdit,
	PptxChartType,
} from '../types/chart';
import { at } from '../../../geometry/indexed';

/** Adds `edit` to the structure edits the save of `chartData` replays, in place. */
export function recordChartStructureEdit(
	chartData: PptxChartData,
	edit: PptxChartStructureEdit,
): void {
	chartData.structureEdits = [...(chartData.structureEdits ?? []), edit];
}

/** `next` (a copy of the chart's data) with `edit` recorded. */
function recorded(next: PptxChartData, edit: PptxChartStructureEdit): PptxChartData {
	recordChartStructureEdit(next, edit);
	return next;
}

// ---------------------------------------------------------------------------
// Chart type compatibility mapping
// ---------------------------------------------------------------------------

/** Chart types that use X/Y numeric data (scatter, bubble). */
const XY_CHART_TYPES = new Set<PptxChartType>(['scatter', 'bubble']);

/**
 * Chart types that support grouping modes (clustered/stacked/percentStacked).
 */
const GROUPING_SUPPORTED_TYPES = new Set<PptxChartType>([
	'bar',
	'line',
	'area',
	'bar3D',
	'line3D',
	'area3D',
]);

// ---------------------------------------------------------------------------
// Add series
// ---------------------------------------------------------------------------

/**
 * Add a new data series to chart data, returning a new `PptxChartData`.
 *
 * If the series' `values` array length does not match the rows of a chart
 * with categories, values are padded with `0` or truncated to match.
 *
 * @param chartData - The current chart data.
 * @param series - The new series to add.
 * @returns A new `PptxChartData` with the series appended.
 *
 * @example
 * ```ts
 * const updated = chartDataAddSeries(chartData, {
 *   name: "Revenue",
 *   values: [100, 200, 300],
 *   color: "#4472C4",
 * });
 * ```
 */
export function chartDataAddSeries(
	chartData: PptxChartData,
	series: { name: string; values: number[]; color?: string },
): PptxChartData {
	const rows = chartData.categories.length > 0 ? chartDataRowCount(chartData) : 0;
	let values = series.values;

	// Align values to the rows of a category-based chart
	if (rows > 0) {
		if (values.length < rows) {
			values = [...values, ...Array(rows - values.length).fill(0)];
		} else if (values.length > rows) {
			values = values.slice(0, rows);
		}
	}

	const newSeries: PptxChartSeries = {
		name: series.name,
		values,
		...(series.color ? { color: series.color } : {}),
	};

	return recorded(
		{
			...chartData,
			series: [...chartData.series, newSeries],
		},
		{ kind: 'add-series' },
	);
}

// ---------------------------------------------------------------------------
// Remove series
// ---------------------------------------------------------------------------

/**
 * Remove a data series by index, returning a new `PptxChartData`.
 *
 * @param chartData - The current chart data.
 * @param seriesIndex - Zero-based index of the series to remove.
 * @returns A new `PptxChartData` without the specified series.
 * @throws {RangeError} If `seriesIndex` is out of bounds.
 *
 * @example
 * ```ts
 * const updated = chartDataRemoveSeries(chartData, 1);
 * ```
 */
export function chartDataRemoveSeries(
	chartData: PptxChartData,
	seriesIndex: number,
): PptxChartData {
	if (seriesIndex < 0 || seriesIndex >= chartData.series.length) {
		throw new RangeError(
			`Series index ${seriesIndex} is out of range. Chart has ${chartData.series.length} series (indices 0\u2013${chartData.series.length - 1}).`,
		);
	}

	return recorded(
		{
			...chartData,
			series: chartData.series.filter((_, i) => i !== seriesIndex),
		},
		{ kind: 'remove-series', series: seriesIndex },
	);
}

// ---------------------------------------------------------------------------
// Update data point
// ---------------------------------------------------------------------------

/**
 * Update a single data point value, returning a new `PptxChartData`.
 *
 * @param chartData - The current chart data.
 * @param seriesIndex - Zero-based series index.
 * @param pointIndex - Zero-based point (category) index.
 * @param value - The new numeric value.
 * @returns A new `PptxChartData` with the updated value.
 * @throws {RangeError} If either index is out of bounds.
 *
 * @example
 * ```ts
 * const updated = chartDataUpdatePoint(chartData, 0, 2, 42);
 * ```
 */
export function chartDataUpdatePoint(
	chartData: PptxChartData,
	seriesIndex: number,
	pointIndex: number,
	value: number,
): PptxChartData {
	if (seriesIndex < 0 || seriesIndex >= chartData.series.length) {
		throw new RangeError(
			`Series index ${seriesIndex} is out of range. Chart has ${chartData.series.length} series.`,
		);
	}
	const series = at(chartData.series, seriesIndex);
	if (pointIndex < 0 || pointIndex >= series.values.length) {
		throw new RangeError(
			`Point index ${pointIndex} is out of range. Series "${series.name}" has ${series.values.length} data points.`,
		);
	}

	const newValues = [...series.values];
	newValues[pointIndex] = value;

	const newSeries = chartData.series.map((s, i) =>
		i === seriesIndex
			? {
					...s,
					values: newValues,
					...(s.blanks?.[pointIndex]
						? { blanks: s.blanks.map((blank, at) => blank && at !== pointIndex) }
						: {}),
				}
			: s,
	);

	return { ...chartData, series: newSeries };
}

// ---------------------------------------------------------------------------
// Change chart type
// ---------------------------------------------------------------------------

/**
 * Change the chart type, adapting grouping and data format as needed.
 *
 * - When switching to a type that does not support grouping, the grouping
 *   field is cleared.
 * - When switching between category-based and XY types, data is preserved
 *   as-is (categories become X labels or vice versa).
 *
 * @param chartData - The current chart data.
 * @param newType - The target chart type.
 * @returns A new `PptxChartData` with the updated type and adapted fields.
 *
 * @example
 * ```ts
 * const lineChart = chartDataChangeType(barChartData, "line");
 * ```
 */
export function chartDataChangeType(
	chartData: PptxChartData,
	newType: PptxChartType,
): PptxChartData {
	const result: PptxChartData = {
		...chartData,
		chartType: newType,
	};

	// Clear grouping if the new type does not support it
	if (!GROUPING_SUPPORTED_TYPES.has(newType)) {
		result.grouping = undefined;
	} else if (!result.grouping) {
		// Set a default grouping for types that support it
		result.grouping = 'clustered';
	}

	// When switching from a category chart to scatter/bubble and categories
	// are non-numeric labels, generate sequential numeric X values
	if (XY_CHART_TYPES.has(newType) && !XY_CHART_TYPES.has(chartData.chartType)) {
		const allCategoriesNumeric = chartData.categories.every((c) => !Number.isNaN(Number(c)));
		if (!allCategoriesNumeric && chartData.categories.length > 0) {
			result.categories = chartData.categories.map((_, i) => String(i + 1));
		}
	}

	return result;
}

// ---------------------------------------------------------------------------
// Add category
// ---------------------------------------------------------------------------

/**
 * Add a new category (data point column) to the chart, in the row after all of its data.
 *
 * Each series gets a `0` in that row; a shorter series is padded with blank points up to it. A
 * chart whose rows have no category labels gets none, as the shared chart operation writes none.
 *
 * @param chartData - The current chart data.
 * @param categoryName - The label for the new category.
 * @returns A new `PptxChartData` with the category appended.
 *
 * @example
 * ```ts
 * const updated = chartDataAddCategory(chartData, "Q4");
 * ```
 */
export function chartDataAddCategory(
	chartData: PptxChartData,
	categoryName: string,
): PptxChartData {
	const row = chartDataRowCount(chartData);
	const padded = <T>(items: readonly T[], fill: T) => [
		...items,
		...Array.from({ length: row - items.length }, () => fill),
	];
	return recorded(
		{
			...chartData,
			categories:
				chartData.categories.length || row === 0
					? [...padded(chartData.categories, ''), categoryName]
					: [],
			series: chartData.series.map((s) => {
				const blanks =
					s.blanks || s.values.length < row
						? [...padded(s.blanks ?? s.values.map(() => false), true), false]
						: undefined;
				return { ...s, values: [...padded(s.values, 0), 0], ...(blanks ? { blanks } : {}) };
			}),
		},
		{ kind: 'add-category' },
	);
}

// ---------------------------------------------------------------------------
// Remove category
// ---------------------------------------------------------------------------

/** The rows of a chart's data: its categories, or the longest series when it has fewer. */
export const chartDataRowCount = (chartData: PptxChartData): number =>
	Math.max(chartData.categories.length, ...chartData.series.map((s) => s.values.length));

/**
 * Remove a category (data point column) by index.
 *
 * The corresponding value, and an X/Y series' X value and bubble size, are also removed
 * from each series.
 *
 * @param chartData - The current chart data.
 * @param categoryIndex - Zero-based index of the category to remove.
 * @param followDataPoint - File > Options > Advanced > "Properties follow
 *   chart data point for current workbook". When `true` (PowerPoint's
 *   default), each series' per-point manual formatting (`c:dPt`, keyed by
 *   `idx`) is re-indexed along with the removed column: the override for the
 *   removed point is dropped, and every override past it shifts down by one
 *   so it stays attached to the same logical point. When `false`, overrides
 *   stay pinned to their original numeric position instead of following the
 *   point they were authored for.
 * @returns A new `PptxChartData` without the specified category.
 * @throws {RangeError} If `categoryIndex` is out of bounds.
 *
 * @example
 * ```ts
 * const updated = chartDataRemoveCategory(chartData, 0);
 * ```
 */
export function chartDataRemoveCategory(
	chartData: PptxChartData,
	categoryIndex: number,
	followDataPoint = true,
): PptxChartData {
	const rows = chartDataRowCount(chartData);
	if (categoryIndex < 0 || categoryIndex >= rows) {
		throw new RangeError(
			`Category index ${categoryIndex} is out of range. Chart has ${rows} categories.`,
		);
	}

	return recorded(
		{
			...chartData,
			categories: chartData.categories.filter((_, i) => i !== categoryIndex),
			series: chartData.series.map((s) => ({
				...s,
				values: s.values.filter((_, i) => i !== categoryIndex),
				...(s.blanks ? { blanks: s.blanks.filter((_, i) => i !== categoryIndex) } : {}),
				...(s.xValues ? { xValues: s.xValues.filter((_, i) => i !== categoryIndex) } : {}),
				...(s.bubbleSizes
					? { bubbleSizes: s.bubbleSizes.filter((_, i) => i !== categoryIndex) }
					: {}),
				...(s.dataPoints && followDataPoint
					? { dataPoints: reindexDataPoints(s.dataPoints, categoryIndex) }
					: {}),
			})),
		},
		{ kind: 'remove-category', category: categoryIndex },
	);
}

/**
 * Re-index per-point formatting overrides (`c:dPt`) after a category at
 * `removedIndex` is deleted: drop the override that pointed at the removed
 * column, and shift every later override's `idx` down by one so it keeps
 * following the same logical data point.
 */
function reindexDataPoints(
	dataPoints: PptxChartDataPoint[],
	removedIndex: number,
): PptxChartDataPoint[] {
	return dataPoints
		.filter((dp) => dp.idx !== removedIndex)
		.map((dp) => (dp.idx > removedIndex ? { ...dp, idx: dp.idx - 1 } : dp));
}
