import { MAX_COL, MAX_ROW, rangeSize } from '../address';
import { getValue } from '../cells';
import { resolveSparklineRange } from '../layout/sparkline-data';
import type { CellValue, ChartObject, Workbook, Worksheet } from '../model';
import { chartSpaceWithSeries } from '../edit/chart-space-series';

/** Maximum cells reread for one chart cache; larger references retain their existing data. */
const MAX_CHART_CACHE_CELLS = 10_000;

/** Reads at most `maxCells` stored range values without recalculating or mutating the workbook. */
function referenceValues(
	workbook: Workbook,
	sheet: Worksheet,
	ref: string | undefined,
	maxCells = MAX_CHART_CACHE_CELLS,
): CellValue[] | undefined {
	if (!ref) return undefined;
	const target = resolveSparklineRange(workbook, sheet, ref);
	if (!target) return undefined;
	const { range } = target;
	const count = rangeSize(range);
	if (
		!Number.isSafeInteger(count) ||
		count < 1 ||
		count > maxCells ||
		range.end.row > MAX_ROW ||
		range.end.col > MAX_COL
	)
		return undefined;
	const values: CellValue[] = [];
	for (let row = range.start.row; row <= range.end.row; row++)
		for (let col = range.start.col; col <= range.end.col; col++)
			values.push(getValue(target.sheet, row, col));
	return values;
}

/** Chart text keeps numeric categories and spells out blanks, logicals and cell errors. */
function categoryValue(value: CellValue): string | number {
	if (value === null) return '';
	if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
	if (typeof value === 'object') return value.error;
	if (typeof value === 'number' && !Number.isFinite(value)) return '';
	return value;
}

/**
 * A serialization-only chart with bounded local-range caches. Oversized or unresolvable
 * references keep their imported caches; series names refresh only from single-cell ranges.
 * Charts carried from the source package never use this helper.
 */
export function chartWithCurrentCaches(
	workbook: Workbook,
	sheet: Worksheet,
	chart: ChartObject,
): ChartObject {
	const cached: ChartObject = {
		...chart,
		series: chart.series.map((series) => {
			const values = referenceValues(workbook, sheet, series.valuesRef);
			const categories = referenceValues(workbook, sheet, series.categoriesRef);
			const name = referenceValues(workbook, sheet, series.nameRef, 1);
			return {
				...series,
				...(name === undefined ? {} : { name: String(categoryValue(name[0] ?? null)) }),
				values:
					values?.map((value) =>
						typeof value === 'number' && Number.isFinite(value) ? value : null,
					) ?? series.values,
				categories: categories?.map(categoryValue) ?? series.categories,
			};
		}),
	};
	return chart.chartSpace
		? { ...cached, chartSpace: chartSpaceWithSeries(chart.chartSpace, chart, cached, true) }
		: cached;
}
