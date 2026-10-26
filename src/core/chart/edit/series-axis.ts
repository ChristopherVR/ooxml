import type { ChartSpace } from '../model';
import { chartSeriesEntryAt, editChartSpace } from './draft';

/** Reassigns an existing coherent axis pair without reallocating ids or reordering series targets. */
export function setChartSeriesAxis(
	space: ChartSpace,
	seriesIndex: number,
	valueAxisId: number | null,
): ChartSpace {
	const { group } = chartSeriesEntryAt(space, seriesIndex);
	if (group.is3D || !['bar', 'line', 'area', 'scatter', 'bubble'].includes(group.kind))
		throw new RangeError('This chart family has no editable axis pair');
	if (new Set(space.plotArea.axes.map((axis) => axis.id)).size !== space.plotArea.axes.length)
		throw new RangeError('Ambiguous chart axis ids');
	const horizontal = group.kind === 'bar' && group.barDirection === 'bar';
	const candidates = space.plotArea.axes.filter(
		(axis) =>
			axis.kind === 'val' &&
			(valueAxisId === null ? axis.position === (horizontal ? 'b' : 'l') : axis.id === valueAxisId),
	);
	if (candidates.length !== 1) throw new RangeError('No unique existing value axis');
	const value = candidates[0]!;
	const category = space.plotArea.axes.find((axis) => axis.id === value.crossAxisId);
	const xy = group.kind === 'scatter' || group.kind === 'bubble';
	if (
		value.id === undefined ||
		category?.id === undefined ||
		category.crossAxisId !== value.id ||
		(xy ? category.kind !== 'val' : !['cat', 'date'].includes(category.kind))
	)
		throw new RangeError('Axis references do not form a coherent pair');
	if (group.axisIds[0] === category.id && group.axisIds[1] === value.id)
		return structuredClone(space);
	if (group.series.length > 1 && group.extLst)
		throw new RangeError('Cannot split an unknown group extension');
	if (group.series.length > 1 && ['stacked', 'percentStacked'].includes(group.grouping ?? ''))
		throw new RangeError('Cannot split a stacked group across axes');
	return editChartSpace(space, (draft) => {
		const { group: current, groupIndex, position } = chartSeriesEntryAt(draft, seriesIndex);
		const selected = {
			...structuredClone(current),
			series: [current.series[position]!],
			axisIds: [category.id!, value.id!],
		};
		const groups = [
			{ ...structuredClone(current), series: current.series.slice(0, position) },
			selected,
			{ ...structuredClone(current), series: current.series.slice(position + 1) },
		].filter((item) => item.series.length);
		draft.plotArea.groups.splice(groupIndex, 1, ...groups);
	});
}
