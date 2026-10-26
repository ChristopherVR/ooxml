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

/** Bounded to existing single-series combo groups, so flattened selections remain stable. */
export function chartSeriesSecondaryAxisSupported(space: ChartSpace, seriesIndex: number): boolean {
	const { group, groupIndex } = chartSeriesEntryAt(space, seriesIndex);
	const first = space.plotArea.groups[0];
	if (!first || groupIndex === 0 || group.series.length !== 1 || group.is3D || first.is3D)
		return false;
	if (![group.kind, first.kind].every((kind) => ['bar', 'line', 'area'].includes(kind)))
		return false;
	if ((group.barDirection === 'bar') !== (first.barDirection === 'bar')) return false;
	return !!axisPair(space, group) && !!axisPair(space, first);
}

export function chartSeriesOnSecondaryAxis(space: ChartSpace, seriesIndex: number): boolean {
	const { group } = chartSeriesEntryAt(space, seriesIndex);
	return group.axisIds[1] !== space.plotArea.groups[0]?.axisIds[1];
}

/** Moves a supported combo group in place; other series, ids, source references and ordering stay intact. */
export function setChartSeriesSecondaryAxis(
	space: ChartSpace,
	seriesIndex: number,
	secondary: boolean,
): ChartSpace {
	if (!chartSeriesSecondaryAxisSupported(space, seriesIndex))
		throw new RangeError('Secondary-axis editing needs an existing single-series 2D combo group');
	return editChartSpace(space, (draft) => {
		const { group } = chartSeriesEntryAt(draft, seriesIndex);
		const primary = draft.plotArea.groups[0]!;
		if (!secondary) {
			group.axisIds = [...primary.axisIds];
			return;
		}
		if (group.axisIds[1] !== primary.axisIds[1]) return;
		const existing = draft.plotArea.groups
			.slice(1)
			.map((candidate) => axisPair(draft, candidate))
			.find((pair) => pair && pair[1].id !== primary.axisIds[1]);
		if (existing) {
			group.axisIds = existing.map((axis) => axis.id!);
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
		const ids: [number, number] = [allocate(), allocate()];
		const horizontal = primary.barDirection === 'bar';
		const axes = defaultChartAxes(horizontal, false, ids);
		axes[0].deleted = true;
		axes[0].position = horizontal ? 'r' : 't';
		axes[1].position = horizontal ? 't' : 'r';
		axes[1].crosses = 'max';
		axes[1].majorGridlines = false;
		draft.plotArea.axes.push(...axes);
		group.axisIds = ids;
	});
}
