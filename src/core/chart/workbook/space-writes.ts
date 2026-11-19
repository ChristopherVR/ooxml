/**
 * The workbook writes that bring an embedded workbook in line with a chart's caches: one per
 * series name, category and value reference (`c:f`), with the values the chart part caches.
 * Literal sources (`c:strLit`, `c:numLit`) and multi-level categories have no single range to
 * write and give none. A blank point is not a cell to write: the cell keeps what it holds (text in
 * a numeric column, nothing at all) unless the edit is what emptied the point.
 *
 * @module chart/workbook/space-writes
 */
import { chartSourceValues } from '../data-cache';
import type { ChartSpace } from '../model';
import type { ChartDataSource } from '../model-series';
import type { ChartWorkbookWrite } from './writer';

/** `values` as cell writes: blanks are skipped, or clear the cell where `before` held a value. */
export const chartWorkbookValues = (
	values: readonly (string | number)[],
	before: readonly (string | number)[],
): (string | null)[] =>
	values.map((value, index) =>
		value !== '' ? String(value) : (before[index] ?? '') === '' ? null : '',
	);

function sourceWrite(
	source: ChartDataSource | undefined,
	before: ChartDataSource | undefined,
): ChartWorkbookWrite | undefined {
	if (!source?.formula || !source.cache || source.kind === 'multiLvlStrRef') return undefined;
	const values = chartWorkbookValues(chartSourceValues(source), chartSourceValues(before));
	const declared = source.cache.pointCount ?? 0;
	return {
		formula: source.formula,
		isNumeric: source.cache.type === 'number',
		values,
		...(declared > values.length ? { cellCount: declared } : {}),
	};
}

/** Every write that makes the workbook cells hold what `space`, edited from `before`, caches. */
export function chartSpaceWorkbookWrites(
	space: ChartSpace,
	before?: ChartSpace,
): ChartWorkbookWrite[] {
	const writes: ChartWorkbookWrite[] = [];
	for (const [g, group] of space.plotArea.groups.entries())
		for (const [s, series] of group.series.entries()) {
			const old = before?.plotArea.groups[g]?.series[s];
			const name = series.tx?.reference;
			if (name?.formula) {
				const text = series.tx?.text ?? chartSourceValues(name).map(String)[0] ?? '';
				writes.push({ formula: name.formula, isNumeric: false, values: [text] });
			}
			for (const key of ['categories', 'values', 'xValues', 'yValues', 'bubbleSizes'] as const) {
				const write = sourceWrite(series[key], old?.[key]);
				if (write) writes.push(write);
			}
		}
	return writes;
}
