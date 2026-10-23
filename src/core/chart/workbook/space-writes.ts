/**
 * The workbook writes that bring an embedded workbook in line with a chart's caches: one per
 * series name, category and value reference (`c:f`), with the values the chart part caches.
 * Literal sources (`c:strLit`, `c:numLit`) and multi-level categories have no single range to
 * write and give none.
 *
 * @module chart/workbook/space-writes
 */
import { chartSourceValues } from '../data-cache';
import type { ChartSpace } from '../model';
import type { ChartDataSource } from '../model-series';
import type { ChartWorkbookWrite } from './writer';

function sourceWrite(source: ChartDataSource | undefined): ChartWorkbookWrite | undefined {
	if (!source?.formula || !source.cache || source.kind === 'multiLvlStrRef') return undefined;
	return {
		formula: source.formula,
		isNumeric: source.cache.type === 'number',
		values: chartSourceValues(source).map(String),
	};
}

/** Every write that makes the workbook cells hold what `space` caches. */
export function chartSpaceWorkbookWrites(space: ChartSpace): ChartWorkbookWrite[] {
	const writes: ChartWorkbookWrite[] = [];
	for (const group of space.plotArea.groups)
		for (const series of group.series) {
			const name = series.tx?.reference;
			if (name?.formula) {
				const text = series.tx?.text ?? chartSourceValues(name).map(String)[0] ?? '';
				writes.push({ formula: name.formula, isNumeric: false, values: [text] });
			}
			for (const source of [
				series.categories,
				series.values,
				series.xValues,
				series.yValues,
				series.bubbleSizes,
			]) {
				const write = sourceWrite(source);
				if (write) writes.push(write);
			}
		}
	return writes;
}
