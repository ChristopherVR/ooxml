import {
	chartSeriesEntryAt,
	setChartPointOptions,
	editChartShapeProperties,
	type ChartPointOptionsEdit,
	type ChartMarker,
} from '../../../chart';
import { fillAppearancePatch } from '../../../chart/edit/appearance-fill';
import type { PptxChartData, PptxChartDataPoint } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';
import { neutralShapeEdit } from './chart-neutral-save-shape';
import { isLineDrawnChartType } from './chart-container-type-map';

function pointEdit(
	edits: NeutralChartEdit[],
	seriesIndex: number,
	old: PptxChartDataPoint | undefined,
	next: PptxChartDataPoint,
): boolean {
	for (const key of ['picture', 'impliedPicture', 'uniqueId'] as const)
		if (!same(old?.[key], next[key])) return false;
	edits.push((space) => {
		const current = chartSeriesEntryAt(space, seriesIndex).series.dataPoints.find(
			(point) => point.index === next.idx,
		);
		const edit: ChartPointOptionsEdit = {};
		for (const key of ['explosion', 'invertIfNegative', 'bubble3D'] as const)
			if (!same(old?.[key], next[key])) Object.assign(edit, { [key]: next[key] ?? null });
		if (!same(old?.spPr, next.spPr))
			edit.spPr = editChartShapeProperties(current?.spPr, neutralShapeEdit(old?.spPr, next.spPr));
		if (!same(old?.gradientFill, next.gradientFill)) {
			const currentShape = edit.spPr ?? current?.spPr;
			edit.spPr = editChartShapeProperties(currentShape, {
				fill:
					fillAppearancePatch(
						currentShape?.fill,
						old?.gradientFill ? { kind: 'gradient', gradient: old.gradientFill } : undefined,
						next.gradientFill
							? { kind: 'gradient', gradient: next.gradientFill }
							: next.spPr?.fillColor === 'none'
								? { kind: 'none' }
								: next.spPr?.fillColor
									? { kind: 'solid', color: next.spPr.fillColor }
									: { kind: 'automatic' },
					) ?? null,
			});
		}
		if (!same(old?.marker, next.marker)) {
			if (!next.marker) edit.marker = null;
			else {
				const marker: ChartMarker = { ...current?.marker, symbol: next.marker.symbol };
				if (next.marker.size === undefined) delete marker.size;
				else marker.size = next.marker.size;
				if (!same(old?.marker?.spPr, next.marker.spPr))
					marker.spPr = editChartShapeProperties(
						current?.marker?.spPr,
						neutralShapeEdit(old?.marker?.spPr, next.marker.spPr),
					);
				edit.marker = marker;
			}
		}
		return setChartPointOptions(space, seriesIndex, next.idx, edit);
	});
	return true;
}

/** Reconciles keyed overrides while retaining extension, theme and relationship tokens. */
export function planNeutralChartPoints(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	probe.series.forEach((series, seriesIndex) => {
		const previous = before.series[seriesIndex];
		if (!previous || same(previous.dataPoints, series.dataPoints)) return;
		const old = previous.dataPoints ?? [],
			next = series.dataPoints ?? [];
		if (
			isLineDrawnChartType(series.seriesChartType ?? probe.chartType) &&
			next.some(
				(point) =>
					!same(old.find((was) => was.idx === point.idx)?.gradientFill, point.gradientFill),
			)
		)
			return;
		if ([old, next].some((list) => new Set(list.map((point) => point.idx)).size !== list.length))
			return;
		const changes: NeutralChartEdit[] = [];
		if (
			!next.every((point) => {
				const was = old.find((value) => value.idx === point.idx);
				return same(was, point) || pointEdit(changes, seriesIndex, was, point);
			})
		)
			return;
		for (const point of old)
			if (!next.some((value) => value.idx === point.idx))
				changes.push((space) => setChartPointOptions(space, seriesIndex, point.idx, null));
		edits.push(...changes);
		series.dataPoints = previous.dataPoints;
	});
}
