// A ChartEx data sheet keeps every hierarchy level instead of flattening the source to one label.
import {
	applyChartDataSheetEdit,
	type ChartDataSheet,
	type ChartDataSheetEdit,
} from '../data-sheet';
import type { ChartExData, ChartExDimension, ChartExLevel, ChartExSpace } from '../ex/model';
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
export function chartExLevelValues(level: ChartExLevel | undefined): string[] {
	const count = Math.max(
		level?.pointCount ?? 0,
		...(level?.points ?? []).map((p) => (p.index ?? 0) + 1),
		0,
	);
	const values = Array<string>(count).fill('');
	for (const point of level?.points ?? [])
		if (point.index !== undefined && point.index >= 0 && point.index < count)
			values[point.index] = point.value;
	return values;
}
/** First-level values for each authored series; parent category levels are exposed separately. */
export function chartExSpaceDataSheet(space: ChartExSpace): ChartExDataSheet {
	const levels = space.plotArea.series.length
		? (chartExSeriesDimension(space, 0, false)?.levels.map(chartExLevelValues) ?? [])
		: [];
	const values = space.plotArea.series.map((_, i) =>
		chartExLevelValues(chartExSeriesDimension(space, i, true)?.levels[0]),
	);
	const count = Math.max(levels[0]?.length ?? 0, ...values.map((v) => v.length), 0);
	const categories = Array.from({ length: count }, (_, i) => levels[0]?.[i] ?? '');
	return {
		categories,
		categoryLevels: levels.length ? levels : [categories],
		series: space.plotArea.series.map((series, i) => ({
			name: series.tx?.value ?? series.tx?.rich?.text ?? '',
			values: Array.from({ length: count }, (_, p) => {
				const v = values[i]?.[p];
				return v !== undefined && v !== '' && Number.isFinite(Number(v)) ? Number(v) : null;
			}),
		})),
	};
}
function rewrite(level: ChartExLevel, values: readonly (string | number | null)[]): void {
	level.pointCount = values.length;
	level.points = values.flatMap((value, index) =>
		value === null || value === '' ? [] : [{ index, value: String(value) }],
	);
}
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
		if (edit.kind === 'series-name') {
			result = renameChartExSeries(result, edit.series, edit.name);
			continue;
		}
		if (edit.kind === 'add-series') {
			result = addChartExSeries(result, next.series.at(-1)!.name);
			continue;
		}
		if (edit.kind === 'remove-series') {
			result = removeChartExSeries(result, edit.series);
			continue;
		}
		result = editChartExSpace(result, (draft) => {
			const seen = new Set<ChartExDimension>();
			for (let i = 0; i < draft.plotArea.series.length; i++) {
				if (edit.kind === 'value') {
					if (i !== edit.series) continue;
					const dimension = chartExSeriesDimension(draft, i, true);
					if (!dimension?.levels[0]) throw new Error('Numeric chart dimension is missing.');
					rewrite(dimension.levels[0], next.series[i]!.values);
					continue;
				}
				for (const numeric of [false, true]) {
					const dimension = chartExSeriesDimension(draft, i, numeric);
					if (!dimension || seen.has(dimension)) continue;
					seen.add(dimension);
					dimension.levels.forEach((level, l) => {
						const values = chartExLevelValues(level);
						if (edit.kind === 'category-label') {
							if (!numeric && l === 0) rewrite(level, next.categories);
						} else if (edit.kind === 'add-category') {
							values.push(numeric ? '0' : '');
							if (!numeric && l === 0) values[values.length - 1] = next.categories.at(-1)!;
							rewrite(level, values);
						} else if (edit.kind === 'remove-category') {
							values.splice(edit.category, 1);
							rewrite(level, values);
						}
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
