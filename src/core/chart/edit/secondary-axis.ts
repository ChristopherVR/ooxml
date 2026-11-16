import type { ChartAxis, ChartPlotGroup, ChartSpace } from '../model';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { defaultChartAxes } from './type-axes';

function axisPair(space: ChartSpace, group: ChartPlotGroup): [ChartAxis, ChartAxis] | undefined {
	if (
		group.axisIds.length !== 2 ||
		new Set(space.plotArea.axes.map((axis) => axis.id)).size !== space.plotArea.axes.length
	)
		return undefined;
	const [x, y] = group.axisIds.map((id) => space.plotArea.axes.find((axis) => axis.id === id));
	return x &&
		y &&
		['cat', 'date'].includes(x.kind) &&
		y.kind === 'val' &&
		x.crossAxisId === y.id &&
		y.crossAxisId === x.id
		? [x, y]
		: undefined;
}

/**
 * The first series anchors the renderer's primary scale; later series of vertical 2D bar, line and
 * area groups can be split in place. A stacked or percent-stacked group gives up its first or last
 * series (each axis then keeps one stack); a series inside the stack is refused.
 */
export function chartSeriesSecondaryAxisSupported(space: ChartSpace, seriesIndex: number): boolean {
	const { group } = chartSeriesEntryAt(space, seriesIndex);
	const first = space.plotArea.groups[0];
	if (!first || seriesIndex === 0 || group.is3D || first.is3D) return false;
	if (![group.kind, first.kind].every((kind) => ['bar', 'line', 'area'].includes(kind)))
		return false;
	// Horizontal combinations are not painted yet.
	if (group.barDirection === 'bar' || first.barDirection === 'bar') return false;
	if (group.series.length > 1) {
		// An unknown group extension cannot be split without guessing what it applies to.
		if (group.extLst) return false;
		// A stacked group keeps one stack per axis: the moved series may leave from either end,
		// but leaving from the middle would part the series around it into two stacks.
		const { position } = chartSeriesEntryAt(space, seriesIndex);
		if (
			(group.grouping === 'stacked' || group.grouping === 'percentStacked') &&
			position > 0 &&
			position < group.series.length - 1
		)
			return false;
	}
	return !!axisPair(space, group) && !!axisPair(space, first);
}

export function chartSeriesOnSecondaryAxis(space: ChartSpace, seriesIndex: number): boolean {
	const { group } = chartSeriesEntryAt(space, seriesIndex);
	return group.axisIds[1] !== space.plotArea.groups[0]?.axisIds[1];
}

/** Split contiguously, retaining flat target positions, authored series ids and group options. */
function isolateSeries(space: ChartSpace, index: number): ChartPlotGroup {
	const { group, groupIndex, position } = chartSeriesEntryAt(space, index);
	if (group.series.length === 1) return group;
	const segments = [
		group.series.slice(0, position),
		[group.series[position]!],
		group.series.slice(position + 1),
	]
		.filter((series) => series.length)
		.map((series) => ({ ...structuredClone(group), series }));
	space.plotArea.groups.splice(groupIndex, 1, ...segments);
	return chartSeriesEntryAt(space, index).group;
}

/** A group's options without its series, to tell whether two neighbours are one group split. */
const groupKey = (group: ChartPlotGroup): string => JSON.stringify({ ...group, series: [] });

/**
 * Joins the group holding series `index` with neighbours of the same type, options and axes, so a
 * series moved back to its axis stacks and clusters with its old group again.
 */
function mergeNeighbours(space: ChartSpace, index: number): void {
	const { groupIndex } = chartSeriesEntryAt(space, index);
	const groups = space.plotArea.groups;
	let at = groupIndex;
	const key = groupKey(groups[at]!);
	if (at > 0 && groupKey(groups[at - 1]!) === key) {
		groups[at - 1]!.series.push(...groups[at]!.series);
		groups.splice(at, 1);
		at--;
	}
	if (at + 1 < groups.length && groupKey(groups[at + 1]!) === key) {
		groups[at]!.series.push(...groups[at + 1]!.series);
		groups.splice(at + 1, 1);
	}
}

export interface ChartSecondaryAxisOptions {
	/**
	 * The ids of a new secondary category and value axis, which must be unused; free ids from
	 * `500000003` by default. Ignored when the series joins axes that exist.
	 */
	axisIds?: readonly [number, number];
}

/** Moves one series, preserving document-order targets and every source data reference. */
export function setChartSeriesSecondaryAxis(
	space: ChartSpace,
	seriesIndex: number,
	secondary: boolean,
	options: ChartSecondaryAxisOptions = {},
): ChartSpace {
	if (!chartSeriesSecondaryAxisSupported(space, seriesIndex))
		throw new RangeError(
			'Secondary-axis editing needs a later series in a supported vertical 2D chart',
		);
	return editChartSpace(space, (draft) => {
		if (chartSeriesOnSecondaryAxis(draft, seriesIndex) === secondary) return;
		const group = isolateSeries(draft, seriesIndex);
		const primary = draft.plotArea.groups[0]!;
		if (!secondary) {
			const left = group.axisIds;
			group.axisIds = [...primary.axisIds];
			mergeNeighbours(draft, seriesIndex);
			// A secondary pair no series plots on any more goes, as Office drops it.
			if (!draft.plotArea.groups.some((other) => other.axisIds.some((id) => left.includes(id))))
				draft.plotArea.axes = draft.plotArea.axes.filter((axis) => !left.includes(axis.id!));
			return;
		}
		if (group.axisIds[1] !== primary.axisIds[1]) return;
		const existing = draft.plotArea.axes
			.filter((axis) => axis.kind === 'cat' || axis.kind === 'date')
			.map((axis) => axisPair(draft, { ...primary, axisIds: [axis.id!, axis.crossAxisId!] }))
			.find((pair) => pair && pair[1].id !== primary.axisIds[1]);
		if (existing) {
			group.axisIds = existing.map((axis) => axis.id!);
			mergeNeighbours(draft, seriesIndex);
			return;
		}
		const used = new Set(draft.plotArea.axes.map((axis) => axis.id));
		const allocate = () => {
			let id = 500000003;
			while (used.has(id)) id++;
			if (id > 4294967295) throw new RangeError('No chart axis IDs remain');
			used.add(id);
			return id;
		};
		const ids: [number, number] = options.axisIds ? [...options.axisIds] : [allocate(), allocate()];
		if (
			options.axisIds &&
			(ids[0] === ids[1] ||
				ids.some((id) => used.has(id) || !Number.isSafeInteger(id) || id < 0 || id > 4294967295))
		)
			throw new RangeError('Secondary axis ids must be two unused whole numbers');
		const horizontal = primary.barDirection === 'bar';
		const axes = defaultChartAxes(horizontal, false, ids);
		axes[0] = { ...structuredClone(axisPair(draft, primary)![0]), id: ids[0], crossAxisId: ids[1] };
		axes[0].deleted = true;
		axes[0].position = horizontal ? 'r' : 't';
		axes[1].position = horizontal ? 't' : 'r';
		axes[1].crosses = 'max';
		axes[1].majorGridlines = false;
		draft.plotArea.axes.push(...axes);
		group.axisIds = ids;
	});
}
