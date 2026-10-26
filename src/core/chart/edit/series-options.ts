import type { ChartSpace, ChartDataPoint, ChartMarker, ChartShapeProperties } from '../model';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { CHART_MARKER_SYMBOLS } from './marker';

export interface ChartSeriesOptionsEdit {
	smooth?: boolean | null;
	invertIfNegative?: boolean | null;
	explosion?: number | null;
	bubble3D?: boolean | null;
	shape?: string | null;
}

function validateOptions(space: ChartSpace, index: number, edit: ChartSeriesOptionsEdit): void {
	const { group } = chartSeriesEntryAt(space, index);
	const families = {
		smooth: ['line', 'scatter', 'stock'],
		invertIfNegative: ['bar', 'bubble'],
		explosion: ['pie', 'doughnut', 'ofPie'],
		bubble3D: ['bubble'],
		shape: ['bar'],
	};
	for (const key of Object.keys(families) as (keyof typeof families)[])
		if (
			edit[key] !== undefined &&
			(!families[key].includes(group.kind) ||
				(key === 'shape' && !group.is3D) ||
				(key === 'smooth' && group.is3D))
		)
			throw new RangeError(`This chart family does not support ${key}`);
	if (
		edit.explosion != null &&
		(!Number.isInteger(edit.explosion) || edit.explosion < 0 || edit.explosion > 400)
	)
		throw new RangeError('Explosion must be an integer from 0 through 400');
	if (
		edit.shape != null &&
		!['box', 'cone', 'coneToMax', 'cylinder', 'pyramid', 'pyramidToMax'].includes(edit.shape)
	)
		throw new RangeError('Unsupported 3D bar shape');
}

/** Partial series options retain every unrelated field and support explicit Automatic resets. */
export function setChartSeriesOptions(
	space: ChartSpace,
	index: number,
	edit: ChartSeriesOptionsEdit,
): ChartSpace {
	validateOptions(space, index, edit);
	return editChartSpace(space, (draft) => {
		const { series } = chartSeriesEntryAt(draft, index);
		for (const key of ['smooth', 'invertIfNegative', 'explosion', 'bubble3D', 'shape'] as const)
			if (edit[key] !== undefined) Object.assign(series, { [key]: edit[key] ?? undefined });
		for (const key of Object.keys(series) as (keyof typeof series)[])
			if (series[key] === undefined) delete series[key];
	});
}

export interface ChartPointOptionsEdit extends Omit<ChartSeriesOptionsEdit, 'shape' | 'smooth'> {
	spPr?: ChartShapeProperties | null;
	marker?: ChartMarker | null;
}

/** Data-point overrides are keyed by authored c:idx, independently of cache or series positions. */
export function setChartPointOptions(
	space: ChartSpace,
	seriesIndex: number,
	pointIndex: number,
	edit: ChartPointOptionsEdit | null,
): ChartSpace {
	if (!Number.isInteger(pointIndex) || pointIndex < 0 || pointIndex > 4294967295)
		throw new RangeError('Invalid chart point index');
	if (edit) {
		validateOptions(space, seriesIndex, edit);
		const marker = edit.marker;
		if (marker) {
			const { group } = chartSeriesEntryAt(space, seriesIndex);
			if (!['line', 'scatter', 'radar', 'stock'].includes(group.kind) || group.is3D)
				throw new RangeError('This chart family has no point markers');
			if (marker.symbol !== undefined && !CHART_MARKER_SYMBOLS.some((s) => s === marker.symbol))
				throw new RangeError('Unsupported marker symbol');
			if (
				marker.size !== undefined &&
				(!Number.isInteger(marker.size) || marker.size < 2 || marker.size > 72)
			)
				throw new RangeError('Invalid marker size');
		}
	}
	return editChartSpace(space, (draft) => {
		const { series } = chartSeriesEntryAt(draft, seriesIndex);
		const position = series.dataPoints.findIndex((point) => point.index === pointIndex);
		if (edit === null) {
			if (position >= 0) series.dataPoints.splice(position, 1);
			return;
		}
		const point: ChartDataPoint = series.dataPoints[position] ?? { index: pointIndex };
		for (const key of ['spPr', 'marker', 'explosion', 'invertIfNegative', 'bubble3D'] as const)
			if (edit[key] === null) delete point[key];
			else if (edit[key] !== undefined) Object.assign(point, { [key]: structuredClone(edit[key]) });
		if (position < 0) series.dataPoints.push(point);
	});
}
