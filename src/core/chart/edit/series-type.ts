import type { ChartSpace } from '../model';
import { chartSeriesEntryAt, editChartSpace } from './draft';
import { setChartType, type ChartTypeTarget } from './type';

/** Changes one category-axis combo series without changing document-order selections or axis ids. */
export function setChartSeriesType(
	space: ChartSpace,
	index: number,
	target: ChartTypeTarget,
): ChartSpace {
	const { group, position } = chartSeriesEntryAt(space, index);
	if (
		group.kind === target.kind &&
		group.is3D === (target.is3D ?? false) &&
		(target.barDirection === undefined || group.barDirection === target.barDirection) &&
		(target.grouping === undefined || group.grouping === target.grouping)
	)
		return structuredClone(space);
	if (
		group.is3D ||
		target.is3D ||
		!['bar', 'line', 'area'].includes(group.kind) ||
		!['bar', 'line', 'area'].includes(target.kind) ||
		group.barDirection === 'bar' ||
		target.barDirection === 'bar'
	)
		throw new RangeError('Series combo types need vertical 2D column, line or area families');
	const pair = group.axisIds.map((id) => space.plotArea.axes.find((axis) => axis.id === id));
	if (
		pair.length !== 2 ||
		!pair[0] ||
		!pair[1] ||
		!['cat', 'date'].includes(pair[0].kind) ||
		pair[1].kind !== 'val' ||
		pair[0].crossAxisId !== pair[1].id ||
		pair[1].crossAxisId !== pair[0].id
	)
		throw new RangeError('Combo types need a coherent category/value axis pair');
	if (group.series.length > 1 && group.extLst)
		throw new RangeError('Cannot split an unknown group extension');
	if (
		position > 0 &&
		position < group.series.length - 1 &&
		['stacked', 'percentStacked'].includes(group.grouping ?? '')
	)
		throw new RangeError('Cannot split the remaining stack into separate groups');
	const isolated = structuredClone(space);
	isolated.plotArea.groups = [
		{ ...structuredClone(group), series: [structuredClone(group.series[position]!)] },
	];
	const converted = setChartType(isolated, target).plotArea.groups[0]!;
	return editChartSpace(space, (draft) => {
		const { group: current, groupIndex, position: at } = chartSeriesEntryAt(draft, index);
		const fragments = [
			{ ...structuredClone(current), series: current.series.slice(0, at) },
			converted,
			{ ...structuredClone(current), series: current.series.slice(at + 1) },
		].filter((part) => part.series.length);
		draft.plotArea.groups.splice(groupIndex, 1, ...fragments);
	});
}
