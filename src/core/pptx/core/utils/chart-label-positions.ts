/**
 * The data label positions a PowerPoint chart model accepts, from the shared
 * table (`chart/label-positions.ts`): a series' labels take what its group
 * accepts, the chart-wide labels only what every group accepts. A series with
 * no type of its own in a combination chart sits in the column group the part
 * starts with.
 *
 * @module utils/chart-label-positions
 */
import { chartGroupElement, chartLabelPositions } from '../../../chart/index';
import type { PptxChartData, PptxChartType } from '../types';
import { NEUTRAL_CHART_TYPE_TARGETS } from './chart-neutral-save-type';

function positionsOfType(data: PptxChartData, type: PptxChartType): readonly string[] | undefined {
	const own = type === 'combo' ? 'bar' : type;
	if (!Object.hasOwn(NEUTRAL_CHART_TYPE_TARGETS, own)) return undefined;
	const target = NEUTRAL_CHART_TYPE_TARGETS[own]!;
	const is3D = own === 'surface' ? data.surfaceTopView === false : (target.is3D ?? false);
	const bar = target.kind === 'bar';
	return chartLabelPositions(
		chartGroupElement(target.kind, is3D),
		bar ? (data.grouping ?? 'clustered') : undefined,
	);
}

/**
 * The positions the labels of series `seriesIndex` (or, without it, the chart-wide labels) accept;
 * `undefined` when the chart type is not a classic one the table describes.
 */
export function chartModelLabelPositions(
	data: PptxChartData,
	seriesIndex?: number,
): readonly string[] | undefined {
	const series = seriesIndex === undefined ? data.series : [data.series[seriesIndex]];
	const types = series.length
		? series.map((item) => item?.seriesChartType ?? data.chartType)
		: [data.chartType];
	let allowed: readonly string[] | undefined;
	for (const type of new Set(types)) {
		const own = positionsOfType(data, type);
		if (!own) return undefined;
		allowed = allowed ? allowed.filter((position) => own.includes(position)) : own;
	}
	return allowed;
}

/** Throws a `RangeError` when the chart model's labels cannot sit at `position`. */
export function assertChartModelLabelPosition(
	data: PptxChartData,
	position: string | null | undefined,
	seriesIndex?: number,
): void {
	if (!position) return;
	const allowed = chartModelLabelPositions(data, seriesIndex);
	if (allowed && !allowed.includes(position))
		throw new RangeError(
			`A ${data.chartType} chart has no data label position "${position}"` +
				(allowed.length ? ` (it takes ${allowed.join(', ')})` : ' (it takes none)'),
		);
}
