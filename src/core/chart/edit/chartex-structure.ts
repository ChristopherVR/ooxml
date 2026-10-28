// Structural ChartEx edits preserve data sets, formatting and unrelated family metadata.
import type { ChartExAxis, ChartExDataLabel, ChartExSeries, ChartExSpace } from '../ex/model';
import { chartExAxisAt, chartExSeriesAt, editChartExSpace } from './chartex';
import { chartExValueType } from './chartex-data';

export const CHART_EX_LAYOUTS = [
	'waterfall',
	'clusteredColumn',
	'boxWhisker',
	'treemap',
	'sunburst',
	'funnel',
	'regionMap',
] as const;
export type ChartExLayout = (typeof CHART_EX_LAYOUTS)[number];

/** Add an independent series and data set. References on the clone are cleared until mapped by its host. */
export function addChartExSeries(
	space: ChartExSpace,
	name: string,
	templateIndex = 0,
): ChartExSpace {
	const template = chartExSeriesAt(space, templateIndex);
	if (!['boxWhisker', 'clusteredColumn'].includes(template.layoutId))
		throw new RangeError('This chart family supports one data series.');
	return editChartExSpace(space, (draft) => {
		const source = chartExSeriesAt(draft, templateIndex);
		const data = source.data ?? draft.chartData.data.find((data) => data.id === source.dataId);
		if (!data) throw new Error('The template series has no data set.');
		const copy = structuredClone(data);
		copy.id = Math.max(-1, ...draft.chartData.data.map((data) => data.id ?? -1)) + 1;
		for (const dimension of copy.dimensions) {
			delete dimension.formula;
			delete dimension.nameFormula;
			if (dimension.kind === 'numDim')
				for (const level of dimension.levels) {
					const count = Math.max(
						level.pointCount ?? 0,
						...level.points.map((point) => (point.index ?? 0) + 1),
						0,
					);
					level.pointCount = count;
					level.points = Array.from({ length: count }, (_, index) => ({ index, value: '0' }));
				}
		}
		draft.chartData.data.push(copy);
		const series: ChartExSeries = {
			...structuredClone(source),
			dataId: copy.id,
			tx: { value: name },
		};
		delete series.data;
		delete series.uniqueId;
		delete series.ownerIndex;
		series.dataPoints = [];
		delete series.dataLabels;
		draft.plotArea.series.push(series);
	});
}

/** Remove a series and its unreferenced data; dependent pareto lines are removed together. */
export function removeChartExSeries(space: ChartExSpace, index: number): ChartExSpace {
	chartExSeriesAt(space, index);
	if (space.plotArea.series.filter((series) => series.ownerIndex === undefined).length < 2)
		throw new RangeError('Cannot remove the last data series.');
	return editChartExSpace(space, (draft) => {
		const old = draft.plotArea.series;
		const removed = new Set([index]);
		old.forEach((series, i) => {
			if (series.ownerIndex === index) removed.add(i);
		});
		const retained = old.map((_, i) => i).filter((i) => !removed.has(i));
		draft.plotArea.series = retained.map((i) => {
			const series = old[i]!;
			if (series.ownerIndex !== undefined) series.ownerIndex = retained.indexOf(series.ownerIndex);
			return series;
		});
		const referenced = new Set(draft.plotArea.series.map((series) => series.dataId));
		const removedIds = new Set([...removed].map((i) => old[i]?.dataId));
		draft.chartData.data = draft.chartData.data.filter(
			(data) => !removedIds.has(data.id) || referenced.has(data.id),
		);
	});
}

/** Change the layout while retaining category levels and the numeric cache and reference. */
export function setChartExLayout(space: ChartExSpace, layout: ChartExLayout): ChartExSpace {
	if (!CHART_EX_LAYOUTS.includes(layout)) throw new RangeError('Unknown ChartEx layout.');
	if (
		layout === 'regionMap' &&
		space.plotArea.series.some((series) => series.layoutId !== 'regionMap')
	)
		throw new RangeError('A region map requires authored geography metadata.');
	const base = space.plotArea.series.filter((series) => series.ownerIndex === undefined);
	if (base.length && base.every((series) => series.layoutId === layout)) return space;
	if (base.length !== 1 && !['boxWhisker', 'clusteredColumn'].includes(layout))
		throw new RangeError('The target layout requires a single data series.');
	return editChartExSpace(space, (draft) => {
		draft.plotArea.series = draft.plotArea.series.filter(
			(series) => series.ownerIndex === undefined,
		);
		for (const series of draft.plotArea.series) {
			const previousType = chartExValueType(series.layoutId);
			const data = series.data ?? draft.chartData.data.find((data) => data.id === series.dataId);
			const value = data?.dimensions.find((d) => d.kind === 'numDim' && d.type === previousType);
			if (value) value.type = chartExValueType(layout);
			series.layoutId = layout;
			delete series.layoutProperties;
			if (layout === 'clusteredColumn') series.layoutProperties = { binning: {} };
		}
		if (['waterfall', 'clusteredColumn', 'boxWhisker'].includes(layout)) {
			draft.plotArea.axes = draft.plotArea.axes.length
				? draft.plotArea.axes
				: [
						{ id: 0, scaling: { kind: 'cat' }, tickLabels: {} },
						{ id: 1, scaling: { kind: 'val' }, tickLabels: {} },
					];
			for (const series of draft.plotArea.series)
				series.axisIds = draft.plotArea.axes.flatMap((axis) =>
					axis.id === undefined ? [] : [axis.id],
				);
		} else {
			draft.plotArea.axes = [];
			for (const series of draft.plotArea.series) series.axisIds = [];
		}
	});
}

/** Patch one axis; unmentioned fields and extension XML stay authored. */
export function setChartExAxis(
	space: ChartExSpace,
	index: number,
	edit: Partial<ChartExAxis>,
): ChartExSpace {
	chartExAxisAt(space, index);
	if (edit.scaling)
		for (const value of Object.values(edit.scaling))
			if (typeof value === 'number' && !Number.isFinite(value))
				throw new RangeError('Axis values must be finite.');
	return editChartExSpace(space, (draft) => {
		Object.assign(chartExAxisAt(draft, index), edit);
	});
}

/** Set or remove a point label override without changing its siblings. */
export function setChartExPointLabel(
	space: ChartExSpace,
	seriesIndex: number,
	pointIndex: number,
	label: Omit<ChartExDataLabel, 'index'> | undefined,
): ChartExSpace {
	if (!Number.isSafeInteger(pointIndex) || pointIndex < 0)
		throw new RangeError('Invalid point index.');
	return editChartExSpace(space, (draft) => {
		const series = chartExSeriesAt(draft, seriesIndex);
		series.dataLabels ??= { labels: [], hidden: [] };
		series.dataLabels.labels = series.dataLabels.labels.filter(
			(label) => label.index !== pointIndex,
		);
		if (label) {
			series.dataLabels.labels.push({ ...label, index: pointIndex });
			series.dataLabels.hidden = series.dataLabels.hidden.filter((index) => index !== pointIndex);
		}
	});
}
