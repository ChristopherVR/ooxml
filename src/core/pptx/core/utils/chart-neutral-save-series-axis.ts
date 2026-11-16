// Format Data Series > Plot Series On through the neutral writer. A move to a secondary value
// axis the part does not have yet (the model's `withChartSeriesOnSecondaryAxis` adds the pair)
// gives the series its own group on a new axis pair, split out of a stack as Office does
// (`setChartSeriesSecondaryAxis`, which Word and Excel use too); a move back merges it into its
// neighbouring group and drops an unused pair. A move onto axes the part already has stays a
// plain reassignment (`setChartSeriesAxis`).
import {
	chartSeriesOnSecondaryAxis,
	chartSeriesSecondaryAxisSupported,
	setChartSeriesAxis,
	setChartSeriesSecondaryAxis,
	type ChartSpace,
} from '../../../chart';
import type { PptxChartData } from '../types';

/** The neutral edit moving series `seriesIndex` to value axis `id` (`null`: the primary one). */
export function seriesAxisEdit(
	after: PptxChartData,
	seriesIndex: number,
	id: number | null,
): (space: ChartSpace) => ChartSpace {
	// Read now: the planner reconciles `after.axes` before the edit runs.
	const category = after.axes?.find(
		(axis) =>
			(axis.axisType === 'catAx' || axis.axisType === 'dateAx') &&
			axis.crossAxisId === id &&
			axis.axisId !== undefined,
	)?.axisId;
	return (space) => {
		const primaryValue = space.plotArea.groups[0]?.axisIds[1];
		const toPrimary = id === null || id === primaryValue;
		const supported = chartSeriesSecondaryAxisSupported(space, seriesIndex);
		if (toPrimary)
			return supported && chartSeriesOnSecondaryAxis(space, seriesIndex)
				? setChartSeriesSecondaryAxis(space, seriesIndex, false)
				: setChartSeriesAxis(space, seriesIndex, id);
		// An axis the part already has: a plain, validated reassignment.
		if (space.plotArea.axes.some((axis) => axis.id === id))
			return setChartSeriesAxis(space, seriesIndex, id);
		// A new pair: keep the model's ids so the saved part and the model agree.
		return setChartSeriesSecondaryAxis(
			space,
			seriesIndex,
			true,
			category === undefined || id === null ? {} : { axisIds: [category, id] },
		);
	};
}

/**
 * Claims series moved onto a secondary axis pair the part does not have yet, or off the last
 * series of one: the pair's two axes are added or dropped by `setChartSeriesSecondaryAxis`, so they
 * are reconciled here, before the axis planner compares axes by position.
 */
export function planNeutralSeriesAxisPair(
	before: PptxChartData,
	probe: PptxChartData,
	edits: ((space: ChartSpace) => ChartSpace)[],
): void {
	if (before.series.length !== probe.series.length) return;
	const was = before.axes ?? [];
	const now = probe.axes ?? [];
	const ids = (list: typeof was) => new Set(list.flatMap((axis) => axis.axisId ?? []));
	const wasIds = ids(was);
	const nowIds = ids(now);
	const added = now.filter((axis) => axis.axisId !== undefined && !wasIds.has(axis.axisId));
	const removed = was.filter((axis) => axis.axisId !== undefined && !nowIds.has(axis.axisId));
	const pair = (list: typeof was) =>
		list.length === 2 &&
		list[0]!.crossAxisId === list[1]!.axisId &&
		list[1]!.crossAxisId === list[0]!.axisId &&
		list.filter((axis) => axis.axisType === 'valAx').length === 1;
	if (added.length && removed.length) return;
	if (added.length ? !pair(added) : removed.length ? !pair(removed) : true) return;
	const pairIds = new Set((added.length ? added : removed).map((axis) => axis.axisId));
	let moved = 0;
	probe.series.forEach((series, index) => {
		const old = before.series[index]!;
		if (old.axisId === series.axisId) return;
		const touches =
			(series.axisId !== undefined && pairIds.has(series.axisId)) ||
			(old.axisId !== undefined && pairIds.has(old.axisId));
		if (!touches) return;
		edits.push(seriesAxisEdit(probe, index, series.axisId ?? null));
		series.axisId = old.axisId;
		moved++;
	});
	if (!moved) return;
	probe.axes = added.length
		? now.filter((axis) => !pairIds.has(axis.axisId))
		: was.map((axis) => now.find((item) => item.axisId === axis.axisId) ?? axis);
}
