// Host-provided live source values fill optional ChartEx caches without depending on a workbook engine.
import type { ChartExDimension, ChartExLevel, ChartExSpace } from './model';

export interface ChartExReferenceValues {
	rows: number;
	columns: number;
	values: readonly (string | number | boolean | null)[];
}
export type ChartExReferenceReader = (formula: string) => ChartExReferenceValues;
function dimensionLevels(dimension: ChartExDimension, read: ChartExReferenceReader): void {
	if (!dimension.formula) return;
	const grid = read(dimension.formula.formula);
	if (grid.rows * grid.columns !== grid.values.length)
		throw new Error('ChartEx source values do not match their range.');
	const horizontal = dimension.formula.direction === 'row';
	const count = horizontal ? grid.columns : grid.rows;
	const length = horizontal ? grid.rows : grid.columns;
	const levels: ChartExLevel[] = Array.from({ length }, (_, level) => {
		const offset = dimension.kind === 'strDim' ? length - 1 - level : level;
		const previous = dimension.levels[level];
		const points = Array.from({ length: count }, (_, index) => {
			const value =
				grid.values[horizontal ? offset * grid.columns + index : index * grid.columns + offset];
			return value === null || value === undefined || value === ''
				? []
				: [{ index, value: String(value) }];
		}).flat();
		return { ...previous, pointCount: count, points };
	});
	dimension.levels = levels;
}
/** A fresh model whose formula caches reflect the host's live source cells. */
export function hydrateChartExSpace(
	space: ChartExSpace,
	read: ChartExReferenceReader,
): ChartExSpace {
	const draft = structuredClone(space);
	for (const data of [
		...draft.chartData.data,
		...draft.plotArea.series.flatMap((series) => (series.data ? [series.data] : [])),
	])
		for (const dimension of data.dimensions) dimensionLevels(dimension, read);
	for (const text of [draft.title?.tx, ...draft.plotArea.series.map((series) => series.tx)])
		if (text?.formula) {
			const grid = read(text.formula.formula);
			if (grid.values.length !== 1) throw new Error('ChartEx text source must be one cell.');
			text.value = String(grid.values[0] ?? '');
		}
	return draft;
}
