// A ChartEx data sheet keeps every hierarchy level instead of flattening the source to one label.
import {
	applyChartDataSheetEdit,
	type ChartDataSheet,
	type ChartDataSheetEdit,
} from '../data-sheet';
import type { ChartExData, ChartExDimension, ChartExLevel, ChartExSpace } from '../ex/model';
import type { ChartCachePoint } from '../model-series';
import {
	chartCacheWithPoint,
	chartCacheWithRow,
	chartCacheWithoutRow,
	chartDensePoints,
} from '../point-limit';
import { chartExSeriesAt, editChartExSpace, renameChartExSeries } from './chartex';
import { chartExValueType } from './chartex-data';
import { addChartExSeries, removeChartExSeries } from './chartex-structure';

export interface ChartExDataSheet extends ChartDataSheet {
	categoryLevels: readonly (readonly string[])[];
}

/** Resolve a series data set without inventing missing or ambiguous source data. */
export function chartExSeriesData(space: ChartExSpace, index: number): ChartExData {
	const series = chartExSeriesAt(space, index);
	if (series.data) return series.data;
	const data = space.chartData.data.filter((data) => data.id === series.dataId);
	if (series.dataId === undefined || data.length !== 1)
		throw new Error('ChartEx data source is missing or ambiguous.');
	return data[0]!;
}
export function chartExSeriesDimension(
	space: ChartExSpace,
	index: number,
	numeric: boolean,
): ChartExDimension | undefined {
	const type = numeric ? chartExValueType(chartExSeriesAt(space, index).layoutId) : 'cat';
	return chartExSeriesData(space, index).dimensions.find(
		(d) =>
			d.kind === (numeric ? 'numDim' : 'strDim') &&
			(d.type === type || (!numeric && d.type === undefined)),
	);
}
export const chartExLevelValues = (level: ChartExLevel | undefined): string[] =>
	chartDensePoints(level ?? { points: [] }, (point) => point.value, '');
/**
 * The plot positions of the series that own data, in sheet order. A series drawn from another one
 * (a pareto line, `@ownerIdx`) has no data of its own and no row in the data sheet.
 */
export function chartExDataSeriesIndices(space: ChartExSpace): number[] {
	return space.plotArea.series.flatMap((series, index) =>
		series.ownerIndex === undefined ? [index] : [],
	);
}
/** First-level values for each authored series; parent category levels are exposed separately. */
export function chartExSpaceDataSheet(space: ChartExSpace): ChartExDataSheet {
	const plot = chartExDataSeriesIndices(space);
	const levels = plot.length
		? (chartExSeriesDimension(space, plot[0]!, false)?.levels.map(chartExLevelValues) ?? [])
		: [];
	const values = plot.map((i) =>
		chartExLevelValues(chartExSeriesDimension(space, i, true)?.levels[0]),
	);
	const count = Math.max(levels[0]?.length ?? 0, ...values.map((v) => v.length), 0);
	const categories = Array.from({ length: count }, (_, i) => levels[0]?.[i] ?? '');
	return {
		categories,
		categoryLevels: levels.length ? levels : [categories],
		series: plot.map((position, i) => {
			const series = space.plotArea.series[position]!;
			return {
				name: series.tx?.value ?? series.tx?.rich?.text ?? '',
				values: Array.from({ length: count }, (_, p) => {
					const v = values[i]?.[p];
					return v !== undefined && v !== '' && Number.isFinite(Number(v)) ? Number(v) : null;
				}),
			};
		}),
	};
}
const levelPoint = (index: number, value: string | number): ChartCachePoint | undefined =>
	value === '' ? undefined : { index, value: String(value) };
/** Apply ordinary grid edits; row edits move all category levels and point overrides together. */
export function applyChartExSpaceDataSheetEdits(
	space: ChartExSpace,
	edits: readonly ChartDataSheetEdit[],
): ChartExSpace {
	let result = space;
	for (const edit of edits) {
		const sheet = chartExSpaceDataSheet(result);
		const next = applyChartDataSheetEdit(sheet, edit);
		if (!next) continue;
		const plot = chartExDataSeriesIndices(result);
		const position = (sheetIndex: number) => {
			const index = plot[sheetIndex];
			if (index === undefined) throw new RangeError(`No chart series at index ${sheetIndex}`);
			return index;
		};
		if (edit.kind === 'series-name') {
			result = renameChartExSeries(result, position(edit.series), edit.name);
			continue;
		}
		if (edit.kind === 'add-series') {
			result = addChartExSeries(result, next.series.at(-1)!.name, plot[0]);
			continue;
		}
		if (edit.kind === 'remove-series') {
			result = removeChartExSeries(result, position(edit.series));
			continue;
		}
		result = editChartExSpace(result, (draft) => {
			const seen = new Set<ChartExDimension>();
			const row = sheet.categories.length;
			for (const [sheetIndex, i] of plot.entries()) {
				if (edit.kind === 'value') {
					if (sheetIndex !== edit.series) continue;
					const levels = chartExSeriesDimension(draft, i, true)?.levels;
					if (!levels?.[0]) throw new Error('Numeric chart dimension is missing.');
					levels[0] = chartCacheWithPoint<ChartCachePoint, ChartExLevel>(
						levels[0],
						edit.category,
						levelPoint(edit.category, edit.value),
					);
					continue;
				}
				for (const numeric of [false, true]) {
					const dimension = chartExSeriesDimension(draft, i, numeric);
					if (!dimension || seen.has(dimension)) continue;
					seen.add(dimension);
					dimension.levels = dimension.levels.map((level, l) => {
						if (edit.kind === 'category-label')
							return numeric || l > 0
								? level
								: chartCacheWithPoint<ChartCachePoint, ChartExLevel>(
										level,
										edit.category,
										levelPoint(edit.category, edit.label),
									);
						if (edit.kind === 'remove-category')
							return chartCacheWithoutRow<ChartCachePoint, ChartExLevel>(level, edit.category);
						const added = numeric ? '0' : l === 0 ? next.categories.at(-1)! : '';
						return chartCacheWithRow<ChartCachePoint, ChartExLevel>(
							level,
							row,
							levelPoint(row, added),
						);
					});
				}
				if (edit.kind === 'remove-category') {
					const series = draft.plotArea.series[i]!;
					const shifted = (index: number) => (index > edit.category ? index - 1 : index);
					series.dataPoints = series.dataPoints
						.filter((p) => p.index !== edit.category)
						.map((p) => ({ ...p, ...(p.index === undefined ? {} : { index: shifted(p.index) }) }));
					if (series.dataLabels) {
						series.dataLabels.labels = series.dataLabels.labels
							.filter((p) => p.index !== edit.category)
							.map((p) => ({
								...p,
								...(p.index === undefined ? {} : { index: shifted(p.index) }),
							}));
						series.dataLabels.hidden = series.dataLabels.hidden
							.filter((p) => p !== edit.category)
							.map(shifted);
					}
					if (series.layoutProperties?.subtotals)
						series.layoutProperties.subtotals = series.layoutProperties.subtotals
							.filter((p) => p !== edit.category)
							.map(shifted);
				}
			}
		});
	}
	return result;
}
